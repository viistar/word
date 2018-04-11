#include "word.h"

bool Word::valid()
{
    return m_bIsValid;
}

void Word::initial()
{
    m_bIsValid = m_wordApplication.setControl("Word.Application");
}

shared_ptr<Document> Word::createDocument(const QString &name)
{
    if(!m_bIsValid)
        return make_shared<Document>();

    QAxObject *documents = m_wordApplication.querySubObject("Documents");
    if(documents)
    {
        QAxObject *document = documents->querySubObject("Add()");
        return make_shared<Document>(document, name);
    }
    else
    {
        return make_shared<Document>();
    }
}

Word *Word::instance()
{
    static Word word;

    return &word;
}

Word::~Word()
{
    if(m_bIsValid)
    {
        m_wordApplication.dynamicCall("Quit()");
    }

}

Word::Word(QObject *parent)
    : QObject(parent),
      m_bIsValid(false)
{
    initial();
}
