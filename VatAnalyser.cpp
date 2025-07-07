#include <QFileInfo>
#include <QDir>

#include <xlsxdocument.h>

#include "../common/utils/CsvReader.h"

#include "VatAnalyser.h"

VatAnalyser::VatAnalyser(const QStringList &csvVatFilePaths)
{
    for (const auto &csvVatFilePath : csvVatFilePaths)
    {
        CsvReader reader{csvVatFilePath,
                    ",",
                    "\""};
        reader.readAll();
        auto dataRode = reader.dataRode();
        int indOrderId = dataRode->header.pos("TRANSACTION_EVENT_ID");
        int indShipmentId = dataRode->header.pos("ACTIVITY_TRANSACTION_ID");
        //int indDateTax = dataRode->header.pos("TAX_CALCULATION_DATE");
        //int indDateTax = dataRode->header.pos("TRANSACTION_DEPART_DATE");
        int indDateTax = dataRode->header.pos("TRANSACTION_COMPLETE_DATE");
        int indUntaxedAmount = dataRode->header.pos("TOTAL_ACTIVITY_VALUE_AMT_VAT_EXCL");
        int indTaxes = dataRode->header.pos("TOTAL_ACTIVITY_VALUE_VAT_AMT");
        int indTaxReportingScheme = dataRode->header.pos("TAX_REPORTING_SCHEME");
        for (const auto &elements : dataRode->lines)
        {
            if (elements.size() > 10)
            {
                const QString &orderId = elements[indOrderId];
                const QString &shipmentId = elements[indShipmentId];
                const QString &untaxedAmount = elements[indUntaxedAmount];
                const QString &taxes = elements[indTaxes];
                const QString &dateTax = formatDateFromVatAmazon(elements[indDateTax]);
                const QString &transactionCreatedId = createTransactionId(orderId, dateTax, untaxedAmount);
                if (transactionCreatedId.contains("408-6395167-2481167"))
                {
                    int TEMP=10;++TEMP;
                }
                const QString &taxReportingScheme = elements[indTaxReportingScheme];
                if (taxReportingScheme == "UNION-OSS")
                {
                    m_orderId_date_amountUntaxed_OSSshipmentId.insert(transactionCreatedId, shipmentId);
                }
                else if (taxReportingScheme == "REGULAR")
                {
                    m_orderId_date_amountUntaxed_REGULARshipmentId.insert(transactionCreatedId, shipmentId);
                }
                else
                {
                    continue;
                }
                bool okDouble = false;
                m_shipmentId_untaxed[shipmentId] = untaxedAmount.toDouble(&okDouble);
                Q_ASSERT(okDouble);
                m_shipmentId_taxes[shipmentId] = taxes.toDouble(&okDouble);
                Q_ASSERT(okDouble);
            }
        }
        //int indTaxes = dataRode->header.pos("TOTAL_ACTIVITY_VALUE_VAT_AMT");
    }
}

QString VatAnalyser::formatDateFromVatAmazon(const QString &dateString) const
{
    QString dateTax;
    if (dateString.indexOf("-") == 2)
    {
        QStringList dateElements = dateString.split("-");
        dateTax += dateElements.takeLast();
        dateTax += "-";
        dateTax += dateElements.takeLast();
        dateTax += "-";
        dateTax += dateElements.takeLast();
    }
    else
    {
        return dateString;
    }
    return dateTax;
}

int VatAnalyser::fixLastCol(const QString &countryCode, int colLast) const
{
    QHash<QString, int> country_max = {
        {"DE", 18}
        , {"IT", 16}
        , {"ES", 29}
    };
    if (country_max.contains(countryCode))
    {
        return qMax(colLast, country_max[countryCode]);
    }
    return colLast;
}

QString VatAnalyser::createTransactionId(
        const QString &orderId, const QString &dateTransaction, const QString &amountHt) const
{
    QString id{orderId};
    static QString sep{"_"};
    id += sep;
    id += dateTransaction;
    id += sep;
    id += amountHt;
    if (!id.contains("."))
    {
        id += ".00";
    }
    else if (id.lastIndexOf(".") == id.size()-2)
    {
        id += "0";
    }
    return id;
}

void VatAnalyser::analyseExcelFile(const QString &excelFilePath) const
{

    QFileInfo excelFileInfo{excelFilePath};
    const QString &baseName = excelFileInfo.baseName();
    const QString &countryCode = getCountryCode(baseName);
    Q_ASSERT(!countryCode.isEmpty());

    QString baseFilePath = QDir{excelFileInfo.path()}.absoluteFilePath(
                baseName);
    const QString &filePathRedErrors = baseFilePath + "-ERRORS.xlsx";
    const QString &filePathCorrected = baseFilePath + "-CORRECTED.xlsx";

    // 1) Open the workbook
    QXlsx::Document xlsx(excelFilePath);
    if (!xlsx.load())
    {
        qWarning() << "Failed to open" << excelFilePath;
        return;
    }

    const QString taxSheet{"Tax return detail"};
    // 2) Switch to the sheet “Tax return detail”
    if (!xlsx.selectSheet(taxSheet))
    {
        qWarning() << "Sheet “Tax return detail” not found in" << excelFilePath;
        return;
    }

    // 3) Determine the used range
    auto dim = xlsx.dimension();  // returns a DocDimension with firstRow, lastRow, firstColumn, lastColumn
    if (!dim.isValid())
    {
        qWarning() << "No data found on sheet Tax return detail.";
        return;
    }
    int rowHeader1   = 2;
    int rowHeader2   = 3;
    int rowFirst = 4;
    int rowLast    = dim.lastRow();
    int colFirst   = dim.firstColumn();
    int colLast    = fixLastCol(countryCode, dim.lastColumn());

    // 4) Read header names from the first row
    QHash<QString, int> col_index;
    for (int col = colFirst; col <= colLast; ++col)
    {
        QVariant v = xlsx.read(rowHeader1, col);
        if (!v.isNull())
        {
            QString header{v.toString().trimmed()};
            col_index[header] = col;
        }
        QVariant v2 = xlsx.read(rowHeader2, col);
        if (!v2.isNull())
        {
            QString header{v2.toString().trimmed()};
            col_index[header] = col;
        }
    }
    int colOrderId = col_index.value("Transaction ID", -1);
    int colDateTransaction = col_index.value("Transaction date", -1);
    const auto &untaxedAmountIndexes = getUntaxedSaleColIndexes(countryCode, col_index);
    const auto &sumTaxIndexes = getColIndexesSumTax(countryCode, col_index);


    QList<int> rowsWrong;
    auto orderId_date_amountUntaxed_OSSshipmentId = m_orderId_date_amountUntaxed_OSSshipmentId;
    auto orderId_date_amountUntaxed_REGULARshipmentId = m_orderId_date_amountUntaxed_REGULARshipmentId;
    int nRowsValid = 0;
    // 5) Iterate every subsequent row, cell by cell
    for (int row = rowFirst + 1; row <= rowLast; ++row)
    {
        const QVariant &cellTransactionId = xlsx.read(row, colOrderId);
        const QVariant &cellTransactionDate = xlsx.read(row, colDateTransaction);
        QVariant cellAmountUntaxed;
        for (const auto &untaxedIndex : untaxedAmountIndexes)
        {
            const QVariant &cell = xlsx.read(row, untaxedIndex);
            if (!cell.isNull())
            {
                if (sumTaxIndexes.size() > 0)
                {
                    double amountTaxes = 0;
                    for (const auto &taxIndex : sumTaxIndexes)
                    {
                        const QVariant &cellTax = xlsx.read(row, taxIndex);
                        if (!cellTax.isNull() && !cellTax.toString().contains("="))
                        {
                            bool taxesIsNumber = false;
                            amountTaxes += cellTax.toDouble(&taxesIsNumber);
                            Q_ASSERT(taxesIsNumber);
                        }
                    }
                    if (qAbs(amountTaxes) < 0.005) // If not vat we ignore the transaction
                    {
                        break;
                    }
                }
                cellAmountUntaxed = cell;
                break;
            }
        }
        if (!cellAmountUntaxed.isNull() && !cellTransactionDate.isNull() && !cellTransactionId.isNull())
        {
            const QString &transactionId = cellTransactionId.toString();
            const QString &transactionDate = cellTransactionDate.toString();
            const QString &transactionAmountUntaxed = cellAmountUntaxed.toString();
            const QString &createdTransactionId = createTransactionId(
                        transactionId, transactionDate, transactionAmountUntaxed);
            if (orderId_date_amountUntaxed_REGULARshipmentId.contains(createdTransactionId))
            {
                orderId_date_amountUntaxed_REGULARshipmentId.take(createdTransactionId);
                ++nRowsValid;
            }
            else if (orderId_date_amountUntaxed_OSSshipmentId.contains(createdTransactionId)) // Error
            {
                rowsWrong << row;
            }
            else
            {
                Q_ASSERT(false); // Unidentified order, we need to understand why
            }
        }

    }
    qInfo() << "Number of rows wrong / valid:" << rowsWrong.size() << "/" << nRowsValid;
    if (rowsWrong.size() > 0)
    {
        QXlsx::Document errorsDoc(excelFilePath);
        errorsDoc.load();
        errorsDoc.selectSheet("Tax return detail");
        errorsDoc.selectSheet(taxSheet);

        QXlsx::Document correctedDoc(excelFilePath);
        correctedDoc.load();
        correctedDoc.selectSheet("Tax return detail");
        correctedDoc.selectSheet(taxSheet);

        // 3) YELLOW FORMAT
        QXlsx::Format yellowFmt;
        //yellowFmt.setPattern(QXlsx::Format::PatternSolid);
        yellowFmt.setPatternBackgroundColor(Qt::yellow);
        for (const auto &row : rowsWrong)
        {
            for (int col = colFirst; col <= colLast; ++col)
            {
                errorsDoc.write(row, col, errorsDoc.read(row, col), yellowFmt);
                correctedDoc.write(row, col, correctedDoc.read(row, col), yellowFmt);
            }
            for (const auto &untaxedIndex : untaxedAmountIndexes)
            {
                if (!correctedDoc.read(row, untaxedIndex).isNull())
                {
                    correctedDoc.write(row, untaxedIndex, QVariant{}, yellowFmt);
                }
            }
        }
        if (!errorsDoc.saveAs(filePathRedErrors))
        {
            qWarning() << "Could not save errors file to" << filePathRedErrors;
        }
        /*
        if (!correctedDoc.saveAs(filePathCorrected))
        {
            qWarning() << "Could not save corrected file to" << filePathCorrected;
        }
        //*/
    }
}

QList<int> VatAnalyser::getUntaxedSaleColIndexes(
        const QString &countryCode, QHash<QString, int> col_index) const
{
    QStringList colNames;
    if (countryCode == "DE")
    {
        colNames << "81";
    }
    else if (countryCode == "IT")
    {
        colNames << "VP2 (NET)";
    }
    else if (countryCode == "ES")
    {
        colNames << "7 (NET)";
        colNames << "14 (NET)";
    }
    else
    {
        Q_ASSERT(false);
    }
    QList<int> indexes;
    for (const auto &colName : colNames)
    {
        Q_ASSERT(col_index.contains(colName));
        indexes << col_index[colName];
    }
    return indexes;
}

QList<int> VatAnalyser::getColIndexesSumTax(
        const QString &countryCode, QHash<QString, int> col_index) const
{
    QStringList colNames;
    if (countryCode == "IT")
    {
        colNames << "VP4 (VAT)";
        colNames << "VP5 (VAT)";
    }
    QList<int> indexes;
    for (const auto &colName : colNames)
    {
        Q_ASSERT(col_index.contains(colName));
        indexes << col_index[colName];
    }
    return indexes;
}

QString VatAnalyser::getCountryCode(const QString &baseName) const
{
    const QStringList &elements = baseName.split("_");
    QString countryCode;
    for (const auto &element : elements)
    {
        static QSet<QString> countryCodes{"DE", "IT", "ES", "PL", "CZ", "UK"};
        if (countryCodes.contains(element))
        {
            return element;
        }
    }
    return QString{};
}

