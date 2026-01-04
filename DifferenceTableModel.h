#ifndef DIFFERENCETABLEMODEL_H
#define DIFFERENCETABLEMODEL_H

#include <QAbstractTableModel>

class DifferenceTableModel : public QAbstractTableModel
{
    Q_OBJECT

public:
    explicit DifferenceTableModel(QObject *parent = nullptr);
    void record(const QString &orderId,
        const QString &shipmentId,
        double untaxedAmount,
        double taxes, const QString &amazon,
        const QString &countryFrom,
        const QString &countryTo,
        const QString &dateTax);
    void removeAllBut(const QSet<QString> &orderIdsToKee);


    // Header:
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;

    // Basic functionality:
    int rowCount(const QModelIndex &parent = QModelIndex()) const override;
    int columnCount(const QModelIndex &parent = QModelIndex()) const override;

    QVariant data(const QModelIndex &index, int role = Qt::DisplayRole) const override;

    Qt::ItemFlags flags(const QModelIndex& index) const override;

private:
    static const QStringList HEADER;
    QList<QVariantList> m_listOfVariantList;
};

#endif // DIFFERENCETABLEMODEL_H
