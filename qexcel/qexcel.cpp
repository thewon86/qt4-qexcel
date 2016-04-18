#include <QAxObject>
#include <QFile>
#include <QStringList>
#include <QDebug>
#include <QDir>

#include "qexcel.h"

QExcel::QExcel() :
    m_ExcelApp(NULL),
    m_workBooks(NULL),
    m_workBook(NULL),
    m_sheets(NULL),
    m_sheet(NULL),
    m_xlsFile(new QString("black.xlsx")),
    m_isOpened(false)
{
    m_ExcelApp = new QAxObject("Excel.Application");
    if(m_ExcelApp != NULL){
        m_isValid = true;
    }
    else{
        m_isValid = false;
        m_isOpened  = false;
    }
}

QExcel::QExcel(const QString &xlsFilePath, QObject *parent) :
    QObject(parent),
    m_ExcelApp(NULL),
    m_workBooks(NULL),
    m_workBook(NULL),
    m_sheets(NULL),
    m_sheet(NULL),
    m_xlsFile(new QString("black.xlsx")),
    m_isOpened(false)
{
    *m_xlsFile = xlsFilePath;
    m_ExcelApp = new QAxObject("Excel.Application");
    if(m_ExcelApp != NULL){
        m_isValid = true;
    }
    else{
        m_isValid = false;
    }
//    m_workBooks = m_ExcelApp->querySubObject("Workbooks");
//    QFile file(xlsFilePath);
//    if (file.exists())
//    {
//        m_workBooks->dynamicCall("Open(const QString&)", xlsFilePath);
//        m_workBook = m_ExcelApp->querySubObject("ActiveWorkBook");
//        m_sheets = m_workBook->querySubObject("WorkSheets");
//    }
//    else{

//        file.open(QIODevice::WriteOnly);
//        file.close();
//    }
}

QExcel::~QExcel()
{
    close();
}

QAxObject *QExcel::getWorkBooks()
{
    return m_workBooks;
}

QAxObject *QExcel::getWorkBook()
{
    return m_workBook;
}

QAxObject *QExcel::getWorkSheets()
{
    return m_sheets;
}

QAxObject *QExcel::getWorkSheet()
{
    return m_sheet;
}

// File
void QExcel::setFileName(const QString &xlsFile)
{
    *m_xlsFile = xlsFile;
}

bool QExcel::open(unsigned int sheet, bool visible)
{
    if(m_isOpened){
        return m_isOpened;
    }
    if(m_xlsFile->isEmpty()){
        m_isOpened = false;
        return false;
    }
    if(m_ExcelApp == NULL){
        m_ExcelApp = new QAxObject("Excel.Application");
        if(m_ExcelApp != NULL){
            m_isValid = true;
        }
        else{
            m_isValid = false;
            return false;
        }
    }
    m_ExcelApp->dynamicCall("SetVisible(bool)", visible);

    m_workBooks = m_ExcelApp->querySubObject("Workbooks");
    QFile xf(*m_xlsFile);
    if(!xf.exists()){
//        creat(m_xlsFile);
        m_workBooks->dynamicCall("Add");
    }
    else{
        m_workBooks->dynamicCall("Open(QString, QVariant)",*m_xlsFile,QVariant(0));
    }
    m_workBook = m_ExcelApp->querySubObject("ActiveWorkBook");
    m_sheets = m_workBook->querySubObject("WorkSheets");

    selectSheet(sheet);
//    saveAs(m_xlsFile);
//    m_sheet = m_sheets->querySubObject("WorkSheet(int)", sheet);
//    if(m_sheet == NULL){
//        addSheet("Sheet5");
//        m_sheet = m_sheets->querySubObject("WorkSheet(int)", 1);
//    }
    m_isOpened = true;
    return true;
}

bool QExcel::open(const QString &lsFile, unsigned int sheet, bool visible)
{
    *m_xlsFile = lsFile;
    return open(sheet, visible);
}

void QExcel::close()
{
    m_ExcelApp->dynamicCall("Quit()");

    if(m_sheet != NULL){
        delete m_sheet;
    }
    delete m_sheets;
    delete m_workBook;
    delete m_workBooks;
    delete m_ExcelApp;

    m_ExcelApp = NULL;
    m_workBooks = NULL;
    m_workBook = NULL;
    m_sheets = NULL;
    m_sheet = NULL;
}

void QExcel::save()
{
    QFile xf(*m_xlsFile);
    if(!xf.exists()){
        saveAs(*m_xlsFile);
    }
    else{
        m_workBook->dynamicCall("Save()");
    }
}

void QExcel::saveAs(const QString &xlsFile)
{
    m_workBook->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",
                            xlsFile,56,QString(""),QString(""),false,false);
}

// Sheet
void QExcel::selectSheet(const QString& sheetName)
{
    m_sheet = m_sheets->querySubObject("Item(const QString&)", sheetName);
}

void QExcel::selectSheet(int index)
{
    m_sheet = m_sheets->querySubObject("Item(int)", index);
}

void QExcel::insertSheet(int index, const QString &sheetName)
{
    int ind = getSheetsCount();
    QString name("");
    m_sheets->querySubObject("Add()");
    if(index == 0){
//        ind = getSheetsCount();
    }
    else{
        if(index < ind){
            ind = index;
        }
    }
    if(!sheetName.isEmpty()){
        name = sheetName;
    }
    else{
        name = QString("Sheet%1").arg(ind);
    }
    setSheetName(ind, name);
//    QAxObject * a = m_sheets->querySubObject("Item(int)", ind);
//    a->setProperty("Name", name);
}

void QExcel::deleteSheet(const QString &sheetName)
{
    QAxObject * a = m_sheets->querySubObject("Item(const QString&)", sheetName);
    if(a != NULL){
        a->dynamicCall("delete");
    }
    else{

    }
}

void QExcel::deleteSheet(int index)
{
    QAxObject * a = m_sheets->querySubObject("Item(int)", index);
    if(a != NULL){
        a->dynamicCall("delete");
    }
    else{

    }
}

void QExcel::setSheetName(int index, const QString &sheetName)
{
    QAxObject * a = m_sheets->querySubObject("Item(int)", index);

    if(a != NULL){
        a->setProperty("Name", sheetName);
    }
    else{

    }
}

void QExcel::setSheetName(const QString &oldName, const QString &newName)
{
    QAxObject * a = m_sheets->querySubObject("Item(const QString&)", oldName);

    if(a != NULL){
        a->setProperty("Name", newName);
    }
    else{

    }
}

QString QExcel::getSheetName()
{
    if(m_sheet != NULL){
        return m_sheet->property("Name").toString();
    }
    else{
        return QString("");
    }
}

QString QExcel::getSheetName(int sheetIndex)
{
    QAxObject * a = m_sheets->querySubObject("Item(int)", sheetIndex);
    if(a != NULL){
        return a->property("Name").toString();
    }
    else{
        return QString("");
    }
}

int QExcel::getSheetsCount()
{
    return m_sheets->property("Count").toInt();
}

void QExcel::setCellString(int row, int column, const QString& value)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Cells(int,int)", row, column);

    if(range != NULL){
        range->dynamicCall("SetValue(const QString&)", value);
    }
    else{
    }
}

void QExcel::setCellFontBold(int row, int column, bool isBold)
{
    QString cell;
    if(m_sheet == NULL){
        return ;
    }
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range = range->querySubObject("Font");
        range->setProperty("Bold", isBold);
    }
    else{
    }
}

void QExcel::setCellFontSize(int row, int column, int size)
{
    QString cell;
    if(m_sheet == NULL){
        return ;
    }
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range = range->querySubObject("Font");
        range->setProperty("Size", size);
    }
    else{
    }
}

void QExcel::mergeCells(const QString& cell)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->setProperty("MergeCells", true);
    }
    else{
    }
}

void QExcel::mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn)
{
    QString cell;
    if(m_sheet == NULL){
        return ;
    }
    cell.append(QChar(topLeftColumn - 1 + 'A'));
    cell.append(QString::number(topLeftRow));
    cell.append(":");
    cell.append(QChar(bottomRightColumn - 1 + 'A'));
    cell.append(QString::number(bottomRightRow));

    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->setProperty("MergeCells", true);
    }
    else{
    }
}

QVariant QExcel::getCellValue(int row, int column)
{
    if(m_sheet == NULL){
        return QVariant();
    }
    QAxObject *range = m_sheet->querySubObject("Cells(int,int)", row, column);
    if(range != NULL){
        return range->property("Value");
    }
    else{
        return QVariant();
    }
}

void QExcel::getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *usedRange = m_sheet->querySubObject("UsedRange");
    if(usedRange != NULL){
        *topLeftRow = usedRange->property("Row").toInt();
        *topLeftColumn = usedRange->property("Column").toInt();
    }
    else{
    }

    QAxObject *rows = usedRange->querySubObject("Rows");
    if(rows != NULL){
        *bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;
    }
    else{
    }

    QAxObject *columns = usedRange->querySubObject("Columns");
    if(columns != NULL){
        *bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
    }
    else{
    }
}

void QExcel::setColumnWidth(int column, int width)
{
    QString columnName;
    if(m_sheet == NULL){
        return ;
    }
    columnName.append(QChar(column - 1 + 'A'));
    columnName.append(":");
    columnName.append(QChar(column - 1 + 'A'));

    QAxObject * col = m_sheet->querySubObject("Columns(const QString&)", columnName);
    if(col != NULL){
        col->setProperty("ColumnWidth", width);
    }
    else{
    }
}

void QExcel::setCellTextCenter(int row, int column)
{
    QString cell;
    if(m_sheet == NULL){
        return ;
    }
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->setProperty("HorizontalAlignment", -4108);//xlCenter
    }
    else{
    }
}

void QExcel::setCellTextWrap(int row, int column, bool isWrap)
{
    QString cell;
    if(m_sheet == NULL){
        return ;
    }
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->setProperty("WrapText", isWrap);
    }
    else{
    }
}

void QExcel::setAutoFitRow(int row)
{
    QString rowsName;
    if(m_sheet == NULL){
        return ;
    }
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    QAxObject * rows = m_sheet->querySubObject("Rows(const QString &)", rowsName);
    if(rows != NULL){
        rows->dynamicCall("AutoFit()");
    }
    else{
    }
}

void QExcel::mergeSerialSameCellsInAColumn(int column, int topRow)
{
    int a,b,c,rowsCount;
    getUsedRange(&a, &b, &rowsCount, &c);

    int aMergeStart = topRow, aMergeEnd = topRow + 1;

    QString value;
    while(aMergeEnd <= rowsCount)
    {
        value = getCellValue(aMergeStart, column).toString();
        while(value == getCellValue(aMergeEnd, column).toString())
        {
            clearCell(aMergeEnd, column);
            aMergeEnd++;
        }
        aMergeEnd--;
        mergeCells(aMergeStart, column, aMergeEnd, column);

        aMergeStart = aMergeEnd + 1;
        aMergeEnd = aMergeStart + 1;
    }
}

void QExcel::clearCell(int row, int column)
{
    QString cell;
    if(m_sheet == NULL){
        return ;
    }
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->dynamicCall("ClearContents()");
    }
    else{
    }
}

void QExcel::clearCell(const QString& cell)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->dynamicCall("ClearContents()");
    }
    else{
    }
}

int QExcel::getUsedRowsCount()
{
    int topRow = 0, bottomRow = 0;
    if(m_sheet == NULL){
        return 0;
    }
    QAxObject *usedRange = m_sheet->querySubObject("UsedRange");
    if(usedRange != NULL){
        topRow = usedRange->property("Row").toInt();
    }
    else{
    }
    QAxObject *rows = usedRange->querySubObject("Rows");
    if(rows != NULL){
        bottomRow = topRow + rows->property("Count").toInt() - 1;
    }
    else{
    }
    return bottomRow;
}

void QExcel::setCellString(const QString& cell, const QString& value)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->dynamicCall("SetValue(const QString&)", value);
    }
    else{
    }
}

void QExcel::setCellFontSize(const QString &cell, int size)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range = range->querySubObject("Font");
        range->setProperty("Size", size);
    }
    else{
    }
}

void QExcel::setCellTextCenter(const QString &cell)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->setProperty("HorizontalAlignment", -4108);//xlCenter
    }
    else{
    }
}

void QExcel::setCellFontBold(const QString &cell, bool isBold)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range = range->querySubObject("Font");
        range->setProperty("Bold", isBold);
    }
    else{
    }
}

void QExcel::setCellTextWrap(const QString &cell, bool isWrap)
{
    if(m_sheet == NULL){
        return ;
    }
    QAxObject *range = m_sheet->querySubObject("Range(const QString&)", cell);
    if(range != NULL){
        range->setProperty("WrapText", isWrap);
    }
    else{
    }
}

void QExcel::setRowHeight(int row, int height)
{
    QString rowsName;
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    if(m_sheet == NULL){
        return ;
    }
    QAxObject *r = m_sheet->querySubObject("Rows(const QString &)", rowsName);
    if(r != NULL){
        r->setProperty("RowHeight", height);
    }
    else{
    }
}
