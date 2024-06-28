#include "tdialogheaders.h"
#include "ui_tdialogheaders.h"
#include    <QMessageBox>

TDialogHeaders::TDialogHeaders(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::TDialogHeaders)
{
    ui->setupUi(this);
    m_model= new QStringListModel(this);
    ui->listView->setModel(m_model); // 显示字符串列表
}

TDialogHeaders::~TDialogHeaders()
{
    delete ui;
}

void TDialogHeaders::setHeaderList(QStringList &headers)
{//设置模型的字符串列表
    m_model->setStringList(headers);
}

QStringList TDialogHeaders::headerList()
{//返回模型的字符串列表
    return  m_model->stringList();
}
