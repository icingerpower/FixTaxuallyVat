#ifndef VATANALYSER_H
#define VATANALYSER_H

#include <QStringList>
#include <QSet>

class VatAnalyser
{
public:
    VatAnalyser(const QStringList &csvVatFilePaths);
    QString createTransactionId(
            const QString &orderId, const QString &dateTransaction, const QString &amountHt) const;
    void analyseExcelFile(const QString &excelFilePath) const;
    QList<int> getUntaxedSaleColIndexes(
            const QString &countryCode, QHash<QString, int> col_index) const;
    QList<int> getColIndexesSumTax(
            const QString &countryCode, QHash<QString, int> col_index) const;
    QString getCountryCode(const QString &baseName) const;
    QString formatDateFromVatAmazon(const QString &date) const;
    int fixLastCol(const QString &countryCode, int lastCol) const;

private:
    QMultiHash<QString, QString> m_orderId_date_amountUntaxed_OSSshipmentId;
    QMultiHash<QString, QString> m_orderId_date_amountUntaxed_REGULARshipmentId;
    QHash<QString, double> m_shipmentId_untaxed;
    QHash<QString, double> m_shipmentId_taxes;
    //QSet<QString> m_shipmentIds_OSS;
    //QSet<QString> m_shipmentIds_REGULAR;
};

#endif // VATANALYSER_H
