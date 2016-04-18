#ifndef QEXCEL_H
#define QEXCEL_H

#include <QString>
#include <QVariant>

class QAxObject;

class QExcel : public QObject
{
public:
    QExcel();
    QExcel(const QString &xlsFilePath, QObject *parent = 0);
    ~QExcel();

public:
    QAxObject * getWorkBooks();
    QAxObject * getWorkBook();
    QAxObject * getWorkSheets();
    QAxObject * getWorkSheet();

public:
    /**************************************************************************/
    /* 文件                                                                   */
    /**************************************************************************/
    void setFileName(const QString &xlsFile);
    bool open(unsigned int sheet = 1, bool visible = false);
    bool open(const QString &lsFile, unsigned int sheet = 1, bool visible = false);
    void close();
    void save();
    void saveAs(const QString &xlsFile);

    /**************************************************************************/
    /* 工作表                                                                 */
    /**************************************************************************/
    void selectSheet(const QString &sheetName);
    void selectSheet(int index);//sheetIndex 起始于 1
    void insertSheet(int index = 0, const QString &sheetName = QString(""));
    void deleteSheet(const QString &sheetName);
    void deleteSheet(int index);
    void setSheetName(int index, const QString &sheetName);
    void setSheetName(const QString &oldName, const QString &newName);
    QString getSheetName();//在 selectSheet() 之后才可调用
    QString getSheetName(int sheetIndex);
    int getSheetsCount();

    /**************************************************************************/
    /* 单元格                                                                 */
    /**************************************************************************/
    void setCellString(int row, int column, const QString &value);
    //cell 例如 "A7"
    void setCellString(const QString &cell, const QString &value);
    //range 例如 "A5:C7"
    void mergeCells(const QString &range);
    void mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn);
    QVariant getCellValue(int row, int column);
    void clearCell(int row, int column);
    void clearCell(const QString &cell);

    /**************************************************************************/
    /* 布局格式                                                               */
    /**************************************************************************/
    void getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn);
    void setColumnWidth(int column, int width);
    void setRowHeight(int row, int height);
    void setCellTextCenter(int row, int column);
    void setCellTextCenter(const QString &cell);
    void setCellTextWrap(int row, int column, bool isWrap);
    void setCellTextWrap(const QString &cell, bool isWrap);
    void setAutoFitRow(int row);
    void mergeSerialSameCellsInAColumn(int column, int topRow);
    int getUsedRowsCount();
    void setCellFontBold(int row, int column, bool isBold);
    void setCellFontBold(const QString &cell, bool isBold);
    void setCellFontSize(int row, int column, int size);
    void setCellFontSize(const QString &cell, int size);

private:
    QAxObject *m_ExcelApp;
    QAxObject *m_workBooks;
    QAxObject *m_workBook;
    QAxObject *m_sheets;
    QAxObject *m_sheet;

    QString *m_xlsFile;     //xls文件路径
    bool m_isOpened;
    bool m_isValid;
    bool m_isVisible;
};

#endif
