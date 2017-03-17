#include "getbillingnum.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    getBillingNum w;
    w.show();

    return a.exec();
}
