#include "DifferenceTableModel.h"

const QStringList DifferenceTableModel::HEADER{
    QObject::tr("Order id")
    , QObject::tr("Shipment ID")
    , QObject::tr("Untaxed amount")
    , QObject::tr("Taxes")
    , QObject::tr("Amazon")
    , QObject::tr("Country from")
    , QObject::tr("Country to")
    , QObject::tr("Date tax")
};

DifferenceTableModel::DifferenceTableModel(QObject *parent)
    : QAbstractTableModel(parent)
{
}

void DifferenceTableModel::record(
        const QString &orderId
        , const QString &shipmentId
        , double untaxedAmount
        , double taxes
        , const QString &amazon
        , const QString &countryFrom
        , const QString &countryTo
        , const QString &dateTax
        )
{
    beginInsertRows(QModelIndex{}, m_listOfVariantList.size(), m_listOfVariantList.size());
    m_listOfVariantList
            << QVariantList{
               orderId
               , shipmentId
               , untaxedAmount
               , taxes
               , amazon
               , countryFrom
               , countryTo
               , dateTax
               };
    endInsertRows();
}

void DifferenceTableModel::removeAllBut(const QSet<QString> &orderIdsToKee)
{
    // Keep only rows whose column 0 ("Order id") is in orderIdsToKee.
    if (m_listOfVariantList.isEmpty())
        return;

    // If nothing to keep -> clear all
    if (orderIdsToKee.isEmpty())
    {
        beginResetModel();
        m_listOfVariantList.clear();
        endResetModel();
        return;
    }

    // Build a filtered list
    QList<QVariantList> kept;
    kept.reserve(m_listOfVariantList.size());

    for (const auto &row : m_listOfVariantList)
    {
        if (row.isEmpty())
            continue;

        const QString orderId = row[0].toString();
        if (orderIdsToKee.contains(orderId))
            kept.push_back(row);
    }

    // If no change, do nothing
    if (kept.size() == m_listOfVariantList.size())
        return;

    beginResetModel();
    m_listOfVariantList = std::move(kept);
    endResetModel();
}


QVariant DifferenceTableModel::headerData(int section, Qt::Orientation orientation, int role) const
{
    if (role == Qt::DisplayRole)
    {
        if (orientation == Qt::Horizontal)
        {
            return HEADER[section];
        }
        else if (orientation == Qt::Vertical)
        {
            return QString::number(section + 1);
        }
    }
    return QVariant{};
}

int DifferenceTableModel::rowCount(const QModelIndex &parent) const
{
    return m_listOfVariantList.size();
}

int DifferenceTableModel::columnCount(const QModelIndex &parent) const
{
    return HEADER.size();
}

QVariant DifferenceTableModel::data(const QModelIndex &index, int role) const
{
    if (role == Qt::DisplayRole || role == Qt::EditRole)
    {
        return m_listOfVariantList[index.row()][index.column()];
    }
    return QVariant();
}

Qt::ItemFlags DifferenceTableModel::flags(const QModelIndex &index) const
{
    return QAbstractItemModel::flags(index) | Qt::ItemIsEditable;
}
