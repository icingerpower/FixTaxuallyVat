#include "MainWindow.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    QCoreApplication::setOrganizationName("Icinger Power");
    QCoreApplication::setApplicationName("Fix Taxually Vat");
    MainWindow w;
    w.show();
    return a.exec();
}
