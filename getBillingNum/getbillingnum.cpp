#include "getbillingnum.h"
#include "ui_getbillingnum.h"

getBillingNum::getBillingNum(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::getBillingNum)
{
    ui->setupUi(this);


    ui->groupBox_message->hide();
    timer = new QTimer(this);
    connect(timer, SIGNAL(timeout()), ui->groupBox_message, SLOT(hide()));

    ui->progressBar_file->hide();
    ui->progressBar_data->hide();

    connect(ui->toolButton_folder, SIGNAL(clicked(bool)), this, SLOT(selectFolder()));
    connect(ui->pushButton_getData, SIGNAL(clicked(bool)), this, SLOT(selectFileName()));

    readSetting();
}

getBillingNum::~getBillingNum()
{
    delete ui;
}

void getBillingNum::closeEvent(QCloseEvent *event)
{
    if (event->isAccepted()){
        writeSetting();
    }
}

void getBillingNum::writeSetting()
{
    QSettings sett("setting.ini", QSettings::IniFormat);
    sett.setValue("size", size());
    sett.setValue("pos", pos());
    sett.setValue("lastPath", folder.absolutePath());
}

void getBillingNum::readSetting()
{
    QSettings sett("setting.ini", QSettings::IniFormat);
    resize(sett.value("size").toSize());
    move(sett.value("pos").toPoint());

    if (!sett.value("lastPath").toString().isEmpty()){

        folder.setPath(sett.value("lastPath").toString());
        ui->lineEdit_dir->setText(folder.toNativeSeparators(sett.value("lastPath").toString()));
        testFolder();
    }
}

void getBillingNum::selectFolder()
{
    QString path = QFileDialog::getExistingDirectory(this, "Выберите папку с файлами...", "HOME", QFileDialog::ShowDirsOnly);
    folder.setPath(path);
    ui->lineEdit_dir->setText(folder.toNativeSeparators(path));

    testFolder();
}

void getBillingNum::selectFileName()
{
    QString _appPath = QApplication::applicationDirPath();
    QString _fname = QFileDialog::getSaveFileName(this, "Save File", "/home/billingNum.xlsx", "Excel (*.xlsx)");
    if (!_fname.isEmpty()){
        get_data(_appPath, _fname);
        make_message("Файл сохранен!", true);
    }
}

void getBillingNum::get_data(QString appPath, QString fname)
{
    report.clear();
    report.append(QString("<p><b>%0</b>: <span style=\" color:#a60000;\">start</span>").arg(QDateTime::currentDateTime().toString("hh:mm dd.MM.yyyy")));
    make_report(report);

    ui->progressBar_file->setValue(0);
    ui->progressBar_file->setVisible(true);
    ui->progressBar_data->setValue(0);
    ui->progressBar_data->setVisible(true);

    QString _temp = QString("%0/temp.xlsx").arg(appPath);
    QFile _file;
    if (_file.copy(QDir::toNativeSeparators(_temp), fname)){
        excel = new QAxObject("Excel.Application");
        excel->setProperty("Visible", 0);
        excel->setProperty("DisplayAlerts", 0);
        wbook = excel->querySubObject("Workbooks");
        book = wbook->querySubObject("Open (const QString&)", fname);
        sheets = book->querySubObject("Sheets");

        QFileInfoList fList = folder.entryInfoList();
        for (int i = 0; i < fList.size(); i++){
            QString fSource = fList.at(i).completeBaseName();

            QString region = fSource.split("_").at(0);
            QString type = fSource.split("_").at(1);

            //проверка наличия нужного листа
            int countList = sheets->dynamicCall("Count()").toInt();
            int tList = 0;
            for (int i = 1; i <= countList; i++){
                currSheet = sheets->querySubObject("Item(Int)", i);
                if (currSheet->dynamicCall("Name").toString() == region){
                    tList = i;                     
                    break;                   
                }
            }
            if (tList == 0){
                //добавить лист
                sheets->dynamicCall("Add()", -4167);
                currSheet = sheets->querySubObject("Item(Int)", 1);
                currSheet->dynamicCall("Name", region);
                //внести данные в 1-2 колонку
                QFile file(fList.at(i).absoluteFilePath());
                if (file.open(QIODevice::ReadOnly)){

                    QDomDocument *doc = new QDomDocument();
                    doc->setContent(&file);
                    QDomNodeList nodeNumList = doc->elementsByTagName("m1");

                    //header
                    QAxObject *hNum = currSheet->querySubObject("Cells(Int, Int)", 1, 1);
                    hNum->dynamicCall("Value", "Номер");
                    QAxObject *hType = currSheet->querySubObject("Cells(Int, Int)", 1, 2);
                    hType->dynamicCall("Value", "Тип");

                    for (int a = 0; a < nodeNumList.count(); a++){
                        QDomElement eNum = nodeNumList.at(a).toElement();
                        QString num = eNum.attribute("u").trimmed();

                        if (!num.isEmpty()){
                            //header
                            QAxObject *dNum = currSheet->querySubObject("Cells(Int, Int)", a + 2, 1);
                            dNum->dynamicCall("Value", num);
                            dNum->dynamicCall("WrapText", 0);
                            QAxObject *dType = currSheet->querySubObject("Cells(Int, Int)", a + 2, 2);
                            dType->dynamicCall("Value", type);
                            dType->dynamicCall("WrapText", 0);
                        }

                        ui->progressBar_data->setValue(qFloor((a + 1) * 100 / nodeNumList.count()));
                        QApplication::processEvents();
                    }
                    //auto_fit
                    QAxObject *aColumns = currSheet->querySubObject("Range(QVariant)", "A1:T5000");
                    aColumns->querySubObject("EntireColumn")->dynamicCall("AutoFit");
                    file.close();

                    report.append(QString("<br /><b>%0</b>: <span style=\" color:#006300;\">%1 - %2</span>").arg(QDateTime::currentDateTime().toString("hh:mm dd.MM.yyyy"))
                                  .arg(fList.at(i).completeBaseName())
                                  .arg(nodeNumList.count()));
                    make_report(report);
                } else {
                    qDebug() << "file don't open";
                }

            } else if (tList > 0){
                currSheet = sheets->querySubObject("Item(Int)", tList);
                // проверить наличие свободных колонок
                int col = 0;
                for (int i = 4; i < 21; i = i + 3){
                    QAxObject *test = currSheet->querySubObject("Cells(Int, Int)", 1, i);
                    QString testStr = test->dynamicCall("Value").toString();
                    if (testStr.isEmpty()){
                        col = i;
                        break;
                    }
                }
                QFile file(fList.at(i).absoluteFilePath());
                if (file.open(QIODevice::ReadOnly)){

                    ui->progressBar_data->setValue(0);

                    QDomDocument *doc = new QDomDocument();
                    doc->setContent(&file);
                    QDomNodeList nodeNumList = doc->elementsByTagName("m1");

                    //header
                    QAxObject *hNum = currSheet->querySubObject("Cells(Int, Int)", 1, col);
                    hNum->dynamicCall("Value", "Номер");
                    QAxObject *hType = currSheet->querySubObject("Cells(Int, Int)", 1, col + 1);
                    hType->dynamicCall("Value", "Тип");

                    for (int a = 0; a < nodeNumList.count(); a++){
                        QDomElement eNum = nodeNumList.at(a).toElement();
                        QString num = eNum.attribute("u").trimmed();

                        if (!num.isEmpty()){
                            //header
                            QAxObject *dNum = currSheet->querySubObject("Cells(Int, Int)", a + 2, col);
                            dNum->dynamicCall("Value", num);
                            dNum->dynamicCall("WrapText", 0);
                            QAxObject *dType = currSheet->querySubObject("Cells(Int, Int)", a + 2, col + 1);
                            dType->dynamicCall("Value", type);
                            dType->dynamicCall("WrapText", 0);
                        }

                        ui->progressBar_data->setValue(qFloor((a + 1) * 100 / nodeNumList.count()));
                        QApplication::processEvents();
                    }
                    //auto_fit
                    QAxObject *aColumns = currSheet->querySubObject("Range(QVariant)", "A1:T5000");
                    //aColumns->dynamicCall("WrapText(Bool)", false);
                    aColumns->querySubObject("EntireColumn")->dynamicCall("AutoFit");
                    file.close();

                    report.append(QString("<br /><b>%0</b>: <span style=\" color:#006300;\">%1 - %2</span>")
                                  .arg(QDateTime::currentDateTime().toString("hh:mm dd.MM.yyyy"))
                                  .arg(fList.at(i).completeBaseName())
                                  .arg(nodeNumList.count()));
                    make_report(report);
                }else {
                    qDebug() << "file don't open";
                }

            }
            ui->progressBar_file->setValue(qFloor((i + 1) * 100 / fList.size()));
            QApplication::processEvents();
        }

        book->dynamicCall("Save()");
        wbook->dynamicCall("Close()");
        excel->dynamicCall("Quit()");
        delete excel;
    }

    report.append(QString("<br /><b>%0</b>: <span style=\" color:#a60000;\">finish</span></p>").arg(QDateTime::currentDateTime().toString("hh:mm dd.MM.yyyy")));
    make_report(report);
    ui->progressBar_file->hide();
    ui->progressBar_data->hide();
}

void getBillingNum::testFolder()
{
    if (folder.isReadable()){
        QStringList nameFilter;
        nameFilter << "*.fp3";
        folder.setNameFilters(nameFilter);
        folder.setFilter(QDir::Files);

        int countFiles = folder.entryInfoList().size();
        if (countFiles == 0){
            make_message("Файлы не найдены", false);
            ui->pushButton_getData->setEnabled(false);
        } else {
            make_message(QString("Найдено файлов: %0").arg(countFiles), true);
            ui->pushButton_getData->setEnabled(true);
        }
    } else {
        make_message("Файлы не найдены", false);
        ui->pushButton_getData->setEnabled(false);
    }
}

void getBillingNum::make_message(QString str, bool v)
{
    ui->groupBox_message->setVisible(true);
    int delay = 0;
    if (v){
        ui->groupBox_message->setStyleSheet("background-color: #006300; border-radius: 9px;");
        delay = 5000;
    } else {
        ui->groupBox_message->setStyleSheet("background-color: #a60000; border-radius: 9px;");
        delay = 15000;
    }
    ui->l_messa->setText(str);
    ui->l_messa->setStyleSheet("color: #FFF5EE; font-weight: bold; ");
    timer->start(delay);
}

void getBillingNum::make_report(QString str)
{
    ui->textEdit_report->clear();
    ui->textEdit_report->setHtml(str);
}
