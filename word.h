#ifndef WORD_H
#define WORD_H

#include <QObject>
#include "document.h"
#include <memory>
#include <ActiveQt/QAxWidget>

using namespace  std;

class Word : public QObject
{
    Q_OBJECT
public:
    bool valid();
    void initial();
    shared_ptr<Document> createDocument(const QString &name);
    static Word* instance();
    ~Word();
private:
    Q_DISABLE_COPY(Word)
    explicit Word(QObject *parent = nullptr);

private:
    QAxWidget m_wordApplication;
    bool m_bIsValid;

};

#endif // WORD_H
