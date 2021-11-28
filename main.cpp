#include <QCoreApplication>
#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>
#include <QClipboard>
#include <QApplication>

#include <iostream>
#include <algorithm>
#include <map>
#include <set>

using namespace std;

/**
 * @brief getDataFromExcel
 *
 * Get data from two columns in Excel and insert it into clipboard
 */
void getDataFromExcel() {
    /**
     * Initializing
     *
     * Getting values:
     *
     * keyColumn   - Column of keys in the table
     * valueColumn - Column of values in the table
     * filePath    - Path to Excel table (Possible to drag and drop file into console to get a file name)
     *
     */

    int keyColumn, valueColumn;

    cout << "Number of a first column with keys: ";
    cin >> keyColumn;

    cout << "Number of a second column with values: ";
    cin >> valueColumn;

    string filePath;

    cout << "Drag and Drop Excel file here" << endl;
    cin >> filePath;

    /**
     * Open Excel
     *
     * Visibility - false
     *
     * Opens Excel table according to filePath variable and gets the first Sheet ftom it
     *
     */
    cout << "Openning Excel" << endl;
    QAxObject* excel     = new QAxObject( "Excel.Application", 0 );
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* workbook  = workbooks->querySubObject( "Open(const QString&)", QString::fromStdString(filePath) );
    QAxObject* sheets    = workbook->querySubObject( "Worksheets" );
    QAxObject* sheet     = sheets->querySubObject( "Item( int )", 1 );

    /**
     * Getting rows count
     */

    QAxObject* usedRange = sheet->querySubObject("UsedRange");
    QAxObject* rows = usedRange->querySubObject("Rows");
    int rowsCount = rows->property("Count").toInt();

    /**
     * Getting data
     *
     * Puts all data into map<string, set<string>>
     */

    cout << "Getting data from excel" << endl;
    map<string, set<string>> resultMap;

    for (int i = 0; i < rowsCount; i++) {
        QAxObject* cell1 = sheet->querySubObject("Cells(QVariant,QVariant)", i + 1, keyColumn);
        QAxObject* cell2 = sheet->querySubObject("Cells(QVariant,QVariant)", i + 1, valueColumn);

        QVariant resultKey   = cell1->property("Value"),
                 resultValue = cell2->property("Value");

        resultMap[resultKey.toString().toStdString()].insert(resultValue.toString().toStdString());

        delete cell1;
        delete cell2;
    }

    /**
     * Putting data into clipboard
     *
     * Transfers map<string, set<string>> into one string and put it into clipboard
     */

    QString clipboardText = "";

    for (const auto &line : resultMap) {
        clipboardText += QString::fromStdString(line.first) + "\t";

        for (const auto &value : line.second) {
            clipboardText += QString::fromStdString(value) + " ";
        }
        clipboardText += "\n";
    }

    QClipboard *clipboard = QApplication::clipboard();
    clipboard->setText(clipboardText, QClipboard::Clipboard);
    cout << "Data copied!" << endl;                           // Done: now one can paste the data into table or somewhere else

    /**
     *   Close excel
     */
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
}

/**
 * @brief execute
 *
 * Procedure to handle errors (Almost no errors are handled)
 */
void execute() {
    cout << "Nashi doblestnie programmisti ne smogli nastroit' russkii yazik, you need to use english instead :)" << endl;

    try {
        getDataFromExcel();
    }  catch (istream::failure e) {
        cout << "How could you break such an easy program))))" << endl;
    }
}

int main(int argc, char *argv[]) {
    QApplication a(argc, argv);

    cin.exceptions(istream::failbit | istream::badbit);

    execute();

    cout << "Everything is done! You can close the window" << endl;

    return a.exec();
}
