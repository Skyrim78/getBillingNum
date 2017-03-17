#include "getbillingnum.h"
#include "ui_getbillingnum.h"

getBillingNum::getBillingNum(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::getBillingNum)
{
    ui->setupUi(this);
    ui->progressBar->hide();
    ui->progressBar_num->hide();

    folder = "/home/ev/Загрузки/testBilling";

    get_data();
}

getBillingNum::~getBillingNum()
{
    delete ui;
}

void getBillingNum::get_data()
{
    for (int r = ui->tableWidget->rowCount() - 1; r >= 0; r--){
        ui->tableWidget->removeRow(r);
    }

    ui->progressBar->setValue(0);
    ui->progressBar->setVisible(true);
    ui->progressBar_num->setValue(0);
    ui->progressBar_num->setVisible(true);

    QStringList nameFilter;
    nameFilter << "*.fp3";

    QDir dir;
    dir.setPath(folder);
    dir.setFilter(QDir::Files);
    dir.setNameFilters(nameFilter);
    QFileInfoList fList = dir.entryInfoList();

    int row = 0;

    for (int i = 0; i < fList.size(); i++){
        QString fname = fList.at(i).completeBaseName();

        QString region = fname.split("_").at(0);
        QString type = fname.split("_").at(1);
        int numCount = 0;

        QFile file(fList.at(i).absoluteFilePath());
        if (file.open(QIODevice::ReadOnly)){

            ui->progressBar_num->setValue(0);

            QDomDocument *doc = new QDomDocument();
            doc->setContent(&file);
            QDomNodeList nodeNumList = doc->elementsByTagName("m1");

            for (int a = 0; a < nodeNumList.count(); a++){
                QDomElement eNum = nodeNumList.at(a).toElement();
                QString num = eNum.attribute("u").trimmed();
                if (!num.isEmpty()){
                    ui->tableWidget->insertRow(row);

                    QTableWidgetItem *itemNum = new QTableWidgetItem();
                    itemNum->setText(num);
                    itemNum->setTextAlignment(Qt::AlignHCenter|Qt::AlignVCenter);
                    ui->tableWidget->setItem(row, 0, itemNum);

                    QTableWidgetItem *itemReg = new QTableWidgetItem();
                    itemReg->setText(region);
                    itemReg->setTextAlignment(Qt::AlignLeft|Qt::AlignVCenter);
                    ui->tableWidget->setItem(row, 1, itemReg);

                    QTableWidgetItem *itemType = new QTableWidgetItem();
                    itemType->setText(type);
                    itemType->setTextAlignment(Qt::AlignHCenter|Qt::AlignVCenter);
                    ui->tableWidget->setItem(row, 2, itemType);

                    numCount++;
                    row++;
                }
                ui->progressBar_num->setValue(qFloor((a + 1) * 100 / nodeNumList.count()));
                QApplication::processEvents();
            }

            file.close();
            qDebug() << region << " - " << type << ": " << numCount;
        }
        ui->progressBar->setValue(qFloor((i + 1) * 100 / fList.count()));
        QApplication::processEvents();
    }
    ui->tableWidget->resizeColumnsToContents();
    ui->tableWidget->horizontalHeader()->setStretchLastSection(true);

    ui->progressBar->hide();
    ui->progressBar_num->hide();
}
