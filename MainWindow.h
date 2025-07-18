#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

public slots:
    void addVatReports();
    void removeVatReports();
    void analyzeResaveTaxuallyFiles();

private:
    Ui::MainWindow *ui;
    void _connectSlots();
};
#endif // MAINWINDOW_H
