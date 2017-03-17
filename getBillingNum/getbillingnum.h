#ifndef GETBILLINGNUM_H
#define GETBILLINGNUM_H

#include <QMainWindow>
#include <QtXml/QtXml>

namespace Ui {
class getBillingNum;
}

class getBillingNum : public QMainWindow
{
    Q_OBJECT

public:
    explicit getBillingNum(QWidget *parent = 0);
    ~getBillingNum();


    QString folder;

public slots:
    void get_data();

private:
    Ui::getBillingNum *ui;
};

#endif // GETBILLINGNUM_H
