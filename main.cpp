#include "mainwindow.h"
#include <QApplication>
#include "word.h"

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();
    auto v = Word::instance();
    v->createDocument("C:/Users/viistar/Documents/xx.docx");
    delete v;
    return a.exec();
}
