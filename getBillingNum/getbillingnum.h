#ifndef GETBILLINGNUM_H
#define GETBILLINGNUM_H

#include <QMainWindow>
#include <QtXml/QtXml>
#include <QFileDialog>
#include <QTimer>
#include <QSettings>
#include <QCloseEvent>

#include <ActiveQt/ActiveQt>


namespace Ui {
class getBillingNum;
}

class getBillingNum : public QMainWindow
{
    Q_OBJECT

public:
    explicit getBillingNum(QWidget *parent = 0);
    ~getBillingNum();

    QTimer *timer;
    QDir folder;

    QAxObject *excel;
    QAxObject *wbook;
    QAxObject *book;
    QAxObject *sheets;
    QAxObject *currSheet;

    QString report;

    virtual void closeEvent(QCloseEvent *event);

public slots:
    void writeSetting();
    void readSetting();

    void selectFolder();
    void selectFileName();

    void get_data(QString appPath, QString fname);
    void testFolder();

    void make_message(QString str, bool v);
    void make_report(QString str);

private:
    Ui::getBillingNum *ui;
};



#endif // GETBILLINGNUM_H
