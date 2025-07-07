#include <QFileDialog>
#include <QSettings>

#include "VatAnalyser.h"

#include "MainWindow.h"
#include "./ui_MainWindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->buttonAnalyseTaxuallyFile->setEnabled(false);
    ui->buttonCopyVatInfos->setEnabled(false);
    _connectSlots();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::_connectSlots()
{
    connect(ui->buttonAddVatReports,
            &QPushButton::clicked,
            this,
            &MainWindow::addVatReports);
    connect(ui->buttonRemoveVatReports,
            &QPushButton::clicked,
            this,
            &MainWindow::removeVatReports);
    connect(ui->buttonAnalyseTaxuallyFile,
            &QPushButton::clicked,
            this,
            &MainWindow::analyzeResaveTaxuallyFiles);
}

void MainWindow::addVatReports()
{
    QSettings settings;
    const QString key{"MainWindow__addVatReports"};
    QString lastDir = settings.value(key, QDir().path()).toString();
    const auto &filePaths = QFileDialog::getOpenFileNames(
                this,
                tr("Amazon VAT report files"),
                lastDir,
                QString{"*.csv"});
    if (!filePaths.isEmpty())
    {
        lastDir = QFileInfo(filePaths[0]).dir().path();
        settings.setValue(key, lastDir);
        ui->listVatReportFiles->addItems(filePaths);
        ui->buttonAnalyseTaxuallyFile->setEnabled(true);
    }
}

void MainWindow::removeVatReports()
{
    auto selItems = ui->listVatReportFiles->selectedItems();

    // Remove each one
    for (auto *item : selItems)
    {
        // takeItem() removes it from the widget and returns ownership
        int row = ui->listVatReportFiles->row(item);
        QListWidgetItem *taken = ui->listVatReportFiles->takeItem(row);

        // delete to free it (unless you have a different ownership plan)
        delete taken;
    }
    if (ui->listVatReportFiles->count() == 0)
    {
        ui->buttonAnalyseTaxuallyFile->setEnabled(false);
    }
}

void MainWindow::analyzeResaveTaxuallyFiles()
{
    QSettings settings;
    const QString key{"MainWindow__analyzeResaveTaxuallyFiles"};
    QString lastDir = settings.value(key, QDir().path()).toString();
    const auto &filePaths = QFileDialog::getOpenFileNames(
                this,
                tr("Amazon VAT report files"),
                lastDir,
                QString{"*.xlsx"});
    QStringList csvVatFilePaths;
    for (int i=0; i< ui->listVatReportFiles->count(); ++i)
    {
        csvVatFilePaths << ui->listVatReportFiles->item(i)->text();
    }
    if (!filePaths.isEmpty())
    {
        lastDir = QFileInfo(filePaths[0]).dir().path();
        settings.setValue(key, lastDir);
        settings.sync();
        VatAnalyser vatAnalyser{csvVatFilePaths};
        for (const auto &filePath : filePaths)
        {
            vatAnalyser.analyseExcelFile(filePath);
        }
    }

    ui->buttonCopyVatInfos->setEnabled(true);
}

