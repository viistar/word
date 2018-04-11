#ifndef DOCUMENT_H
#define DOCUMENT_H

#include <QAxObject>
#include <QObject>

class Document : public QObject
{
    Q_OBJECT
public:
    explicit Document(QObject *parent = nullptr);
    explicit Document(QAxObject *doc, const QString &name, QObject *parent = nullptr);
    ~Document();
    void setDocument(QAxObject *doc, const QString &name);
    void save();
private:
    QAxObject* m_wordDocument;
    QString m_strDocumentName;
    bool m_bIsSave;
};

#endif // DOCUMENT_H
