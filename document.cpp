#include "document.h"

Document::Document(QObject *parent)
    : QObject(parent),
      m_wordDocument(nullptr),
      m_strDocumentName(QString()),
      m_bIsSave(false)
{

}

Document::Document(QAxObject *doc, const QString &name, QObject *parent)
    : QObject(parent),
      m_bIsSave(false)
{
    setDocument(doc, name);
}

Document::~Document()
{
    save();
}

void Document::setDocument(QAxObject *doc, const QString &name)
{
    m_wordDocument = doc;
    m_strDocumentName = name;
}

void Document::save()
{
    if(m_wordDocument && (!m_bIsSave))
    {
        m_wordDocument->dynamicCall("SaveAs(const QString &)", m_strDocumentName);
    }
    else if(m_wordDocument)
    {
        m_wordDocument->dynamicCall("Save()");
    }
    else
    {

    }
}
