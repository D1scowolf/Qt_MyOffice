#include "tformtable.h"
#include "ui_tformtable.h"

#include    <QToolBar>
#include    <QVBoxLayout>
#include    <QFileDialog>
#include    <QFile>
#include    <QSaveFile>
#include    <QTextStream>
#include    <QFontDialog>
#include    <QFileInfo>
#include    <QMessageBox>

TFormTable::TFormTable(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::TFormTable)
{
    ui->setupUi(this);

//    // 自行设计ToolBar
//    QToolBar* locToolBar = new QToolBar("表格",this);
//    locToolBar->addAction(ui->actOpenExcel);
//    locToolBar->addAction(ui->actSaveExcel);
//    locToolBar->addSeparator();
//    locToolBar->addAction(ui->actCut);
//    locToolBar->addAction(ui->actCopy);
//    locToolBar->addAction(ui->actPaste);
//    locToolBar->addAction(ui->actUndo);
//    locToolBar->addAction(ui->actRedo);
//    locToolBar->addSeparator();
//    locToolBar->addAction(ui->actClose);
//    addToolBarBreak();
//    locToolBar->addAction(ui->actSize);
//    locToolBar->addAction(ui->actSetHeader);
//    locToolBar->addSeparator();
//    locToolBar->addAction(ui->actBold);
//    locToolBar->addAction(ui->actItalic);
//    locToolBar->addAction(ui->actUnder);
//    locToolBar->addSeparator();
//    locToolBar->addAction(ui->actAlignLeft);
//    locToolBar->addAction(ui->actAlignCenter);
//    locToolBar->addAction(ui->actAlignRight);

//    locToolBar->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);

    m_model = new QStandardItemModel(12, 6, this);
    m_selection = new QItemSelectionModel(m_model,this);

    // 当前单元格变化时的信号与槽
    connect(m_selection, &QItemSelectionModel::currentChanged,
            this, &TFormTable::do_currentChanged);

    // 为tableView设置数据模型
    ui->tableView->setModel(m_model);
    ui->tableView->setSelectionModel(m_selection);
    // 操作多样性——shift和ctrl功能在库里就已经实现
    ui->tableView->setSelectionMode(QAbstractItemView::ExtendedSelection);
    ui->tableView->setSelectionBehavior(QAbstractItemView::SelectItems);

    // 状态栏的实现
    labCellPos = new QLabel("当前单元格：",this);
    labCellPos->setMinimumWidth(180);
    labCellPos->setAlignment(Qt::AlignHCenter);

    labCellText = new QLabel("单元格内容：",this);
    labCellText->setMinimumWidth(150);

    ui->statusBar->addWidget(labCellPos);
    ui->statusBar->addWidget(labCellText);
    //    QMessageBox::information(this, "消息", "表格窗口被创建");
}

TFormTable::~TFormTable()
{
    //    QMessageBox::information(this, "消息", "FormTable窗口被删除和释放");
    delete ui;
}

// 选择单元格变化时的响应
void TFormTable::do_currentChanged(const QModelIndex &current, const QModelIndex &previous)
{
    Q_UNUSED(previous);

    if (current.isValid())
    {
        labCellPos->setText(QString::asprintf("当前单元格：%d行，%d列", current.row() + 1, current.column() + 1));

        QStandardItem *aItem = m_model->itemFromIndex(current);
        labCellText->setText("单元格内容："+aItem->text());

        QFont font = aItem->font();
        ui->actBold->setChecked(font.bold());
    }
}

// 左对齐
void TFormTable::on_actAlignLeft_triggered()
{
    // 检查有无可选择的项
    if(!m_selection->hasSelection())
        return;

    // 获取选择的单元格的模型索引列表，可多选
    QModelIndexList IndexList = m_selection->selectedIndexes();

    for(int i = 0; i < IndexList.count(); ++i)
    {
        // 获取其中一个模型索引
        QModelIndex aIndex = IndexList.at(i);
        // 获取一个单元格的对象
        QStandardItem *aItem = m_model->itemFromIndex(aIndex);
        aItem->setTextAlignment(Qt::AlignLeft | Qt::AlignVCenter);
    }
}

// 居中对齐
void TFormTable::on_actAlignCenter_triggered()
{
    // 检查有无可选择的项
    if(!m_selection->hasSelection())
        return;

    // 获取选择的单元格的模型索引列表，可多选
    QModelIndexList IndexList = m_selection->selectedIndexes();

    for(int i = 0; i < IndexList.count(); ++i)
    {
        // 获取其中一个模型索引
        QModelIndex aIndex = IndexList.at(i);
        // 获取一个单元格的对象
        QStandardItem *aItem = m_model->itemFromIndex(aIndex);
        aItem->setTextAlignment(Qt::AlignHCenter | Qt::AlignVCenter);
    }
}

// 右对齐
void TFormTable::on_actAlignRight_triggered()
{
    // 检查有无可选择的项
    if(!m_selection->hasSelection())
        return;

    // 获取选择的单元格的模型索引列表，可多选
    QModelIndexList IndexList = m_selection->selectedIndexes();

    for(int i = 0; i < IndexList.count(); ++i)
    {
        // 获取其中一个模型索引
        QModelIndex aIndex = IndexList.at(i);
        // 获取一个单元格的对象
        QStandardItem *aItem = m_model->itemFromIndex(aIndex);
        aItem->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
    }
}

// 设置字体为粗体
void TFormTable::on_actBold_triggered(bool checked)
{
    if(!m_selection->hasSelection())
        return;

    QModelIndexList selectedIndex = m_selection->selectedIndexes();

    for(int i = 0; i < selectedIndex.count(); ++i)
    {
        QModelIndex aIndex = selectedIndex.at(i);
        QStandardItem *aItem = m_model->itemFromIndex(aIndex);
        QFont font = aItem->font();
        font.setBold(checked);
        aItem->setFont(font);
    }
}

// 设置字体带下划线
void TFormTable::on_actUnder_triggered(bool checked)
{
    if(!m_selection->hasSelection())
        return;

    QModelIndexList selectedIndex = m_selection->selectedIndexes();

    for(int i = 0; i < selectedIndex.count(); ++i)
    {
        QModelIndex aIndex = selectedIndex.at(i);
        QStandardItem *aItem = m_model->itemFromIndex(aIndex);
        QFont font = aItem->font();
        font.setUnderline(checked);
        aItem->setFont(font);
    }
}

// 设置字体为斜体
void TFormTable::on_actItalic_triggered(bool checked)
{
    if(!m_selection->hasSelection())
        return;

    QModelIndexList selectedIndex = m_selection->selectedIndexes();

    for(int i = 0; i < selectedIndex.count(); ++i)
    {
        QModelIndex aIndex = selectedIndex.at(i);
        QStandardItem *aItem = m_model->itemFromIndex(aIndex);
        QFont font = aItem->font();
        font.setItalic(checked);
        aItem->setFont(font);
    }
}

// 打开Excel
void TFormTable::on_actOpenExcel_triggered()
{
    // 打开的时候重置
    m_model = new QStandardItemModel(0, 0, this);
    ui->tableView->setModel(m_model);

    // 同之前的QFileDialog打开
    QString curPath = QDir::currentPath();
    QString dlgTitle = "打开一个Excel文件";
    QString filter = "Excel 文件(*.xls)";
    QString aFileName = QFileDialog::getOpenFileName(this, dlgTitle, curPath, filter);

    if (aFileName.isEmpty())
        return;

    QFileInfo fileinfo(aFileName);
    QDir::setCurrent(fileinfo.absolutePath());

    QFile file(aFileName);
    if (file.open(QIODevice::ReadOnly))
    {
        QTextStream in(&file);

        QString line = in.readLine();
        QStringList header = line.split("\t");
        m_model->setHorizontalHeaderLabels(header);

        while (!in.atEnd())
        {
            line = in.readLine();
            QStringList data = line.split("\t");
            QList<QStandardItem*> items;
            for (int i = 0; i < data.size(); i++)
            {
                items.append(new QStandardItem(data.at(i)));
            }
            m_model->appendRow(items);
        }

        file.close();
    }

    m_selection = new QItemSelectionModel(m_model,this);
    ui->tableView->setSelectionModel(m_selection);
    ui->tableView->setSelectionMode(QAbstractItemView::ExtendedSelection);
    ui->tableView->setSelectionBehavior(QAbstractItemView::SelectItems);

    connect(m_selection, &QItemSelectionModel::currentChanged,
            this, &TFormTable::do_currentChanged);

    //QMessageBox::information(this, "提示", "导入完成");
}

// 保存为Excel
void TFormTable::on_actSaveExcel_triggered()
{
    // QStandardItemModel获取当前QTableView里的内容
    // 事实上我定义了相关变量
    QStandardItemModel *model = static_cast<QStandardItemModel*>(ui->tableView->model());

    // 与之前文本文档操作一致，设置路径
    QString curPath = QDir::currentPath();
    QString dlgTitle = "保存为一个Excel文件";
    QString filter = "Excel 文件(*.xls)";
    QString aFileName = QFileDialog::getSaveFileName(this, dlgTitle, curPath, filter);

    if (aFileName.isEmpty())
        return;

    QFileInfo fileinfo(aFileName);
    QDir::setCurrent(fileinfo.absolutePath());

    QFile file(aFileName);
    if (file.open(QIODevice::WriteOnly | QIODevice::Truncate))
    {
        // 用QTextStream写入
        QTextStream out(&file);

        // 写入的方式就是表头+表内容，用\t \n分隔
        for (int column = 0; column < model->columnCount(); ++column)
        {
            out<<model->headerData(column, Qt::Horizontal).toString() << "\t";
        }
        out << "\n";

        for (int row = 0; row < model->rowCount(); ++row)
        {
            for (int column = 0; column < model->columnCount(); ++column)
            {
                QModelIndex index = model->index(row, column);
                out << model->data(index).toString() << "\t";
            }
            out<<"\n";
        }

        file.close();
    }

    //QMessageBox::information(this, "提示", "导出完成");
}

// 添加行
void TFormTable::on_actAppend_triggered()
{
    // QList连续输出
    QList<QStandardItem*> aItemList;
    QStandardItem *aItem;
    int len = m_model->columnCount();

    for (int i = 0; i < len; ++i)
    {
        // 添加的每一格都是空白的
        aItem = new QStandardItem();
        aItemList << aItem;
    }

    m_model->insertRow(m_model->rowCount(), aItemList);
    // 最后一行、第一列的索引
    QModelIndex curIndex = m_model->index(m_model->rowCount() - 1, 0);
    m_selection->clearSelection();
    // 设置当前选择行
    m_selection->setCurrentIndex(curIndex, QItemSelectionModel::Select);
}

// 插入行
void TFormTable::on_actInsert_triggered()
{
    // QList连续输出
    QList<QStandardItem*> aItemList;
    QStandardItem *aItem;
    int len = m_model->columnCount();

    for (int i = 0; i < len; ++i)
    {
        // 添加的每一格都是空白的
        aItem = new QStandardItem();
        aItemList << aItem;
    }

    QModelIndex curIndex=m_selection->currentIndex();
    // 在当前选择行前面插入
    m_model->insertRow(curIndex.row(),aItemList);
    m_selection->clearSelection();
    // 设置当前选择行
    m_selection->setCurrentIndex(curIndex, QItemSelectionModel::Select);
}

// 删除行
void TFormTable::on_actDelete_triggered()
{
    // 删除当前选择行
    QModelIndex curIndex=m_selection->currentIndex();

    if (curIndex.row() == m_model->rowCount() - 1)
        m_model->removeRow(curIndex.row());
    else
    {
        m_model->removeRow(curIndex.row());
        m_selection->setCurrentIndex(curIndex,QItemSelectionModel::Select);
    }
}

// 定义表格大小
void TFormTable::on_actSize_triggered()
{
    TDialogSize *dlgTableSize=new TDialogSize(this);
    // 设置对话框固定大小,防止用户误操作
    dlgTableSize->setWindowFlag(Qt::MSWindowsFixedSizeDialogHint);
    dlgTableSize->setRowColumn(m_model->rowCount(),m_model->columnCount());

    int ret = dlgTableSize->exec();   // 以模态方式显示对话框
    if (ret == QDialog::Accepted)     // OK键被按下，获取对话框上的输入，设置行数和列数
    {
        int cols = dlgTableSize->columnCount();
        m_model->setColumnCount(cols);
        int rows = dlgTableSize->rowCount();
        m_model->setRowCount(rows);
    }
    delete dlgTableSize; //删除对话框
}

// 设置表头
void TFormTable::on_actSetHeader_triggered()
{
    if (dlgSetHeaders == nullptr) // 如果对象没有被创建过，就创建对象
        dlgSetHeaders = new TDialogHeaders(this);

    if (dlgSetHeaders->headerList().count() != m_model->columnCount())
    {
        QStringList strList;
        // 获取现有的表头标题
        for (int i = 0; i < m_model->columnCount(); i++)
            strList.append(m_model->headerData(i,Qt::Horizontal,Qt::DisplayRole).toString());
        dlgSetHeaders->setHeaderList(strList);
    }

    int ret=dlgSetHeaders->exec(); // 以模态方式显示对话框

    if (ret==QDialog::Accepted)
    {
        QStringList strList=dlgSetHeaders->headerList();
        m_model->setHorizontalHeaderLabels(strList);
    }
}
