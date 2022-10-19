#include "pch.h"
#include "Classes.h"
#include "Header.h"

using namespace Microsoft::Office::Interop::Excel;
using namespace System::Runtime::InteropServices;

Graph g;


void DisplayMenu() {
    cout << "\t Меню" << endl;
    cout << "1 - Алгоритм поиска в ширину" << endl;
    cout << "2 - Добавить новую вершину" << endl;
    cout << "3 - Добавить новый путь" << endl;
    cout << "4 - Отобразить таблицу смежности" << endl;
    cout << endl;
    cout << "Esc - выход из программы" << endl;
}

void ExcelAddVertexs() {
    String^ fileName;
    OpenFileDialog^ openDlg = gcnew OpenFileDialog();

    openDlg->Multiselect = false;
    openDlg->Filter = "Excel Document(*.xlsx;*.xls)|*.xlsx; *.xls";
    openDlg->Title = "Открыть файл с вершинами...";

    if (System::Windows::Forms::DialogResult::OK == openDlg->ShowDialog())
    {
        fileName = openDlg->FileName;

        Microsoft::Office::Interop::Excel::Application^ ExcelApplication = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();

        Workbook^ wb = ExcelApplication->Workbooks->Open(fileName, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
        Worksheet^ exWs = safe_cast<Worksheet^>(ExcelApplication->ActiveSheet);

        // Найти последнюю заполненую строку
        int lastUsedRow = exWs->Cells->Find("*", System::Reflection::Missing::Value,
            System::Reflection::Missing::Value, System::Reflection::Missing::Value,
            Microsoft::Office::Interop::Excel::XlSearchOrder::xlByRows, Microsoft::Office::Interop::Excel::XlSearchDirection::xlPrevious,
            false, System::Reflection::Missing::Value, System::Reflection::Missing::Value)->Row;

        vector<string> ArrayTableStr;
        String^ tmp;
        string buff;
        for (int i = 1; i <= lastUsedRow; i++) {
            tmp = ((Microsoft::Office::Interop::Excel::Range^)exWs->Cells[(System::Object^)i, (System::Object^)1])->Value2->ToString();
            buff = msclr::interop::marshal_as<std::string>(tmp);
            ArrayTableStr.push_back(buff);
        }
        for (int i = 0; i < lastUsedRow; i++) {
            g.CreateVertex(ArrayTableStr[i]);
        }
        ExcelApplication->Workbooks->Close();
    }
    else
        exit(0);
}

void ExcelAddWay() {
    String^ fileName;
    OpenFileDialog^ openDlg = gcnew OpenFileDialog();

    openDlg->Multiselect = false;
    openDlg->Filter = "Excel Document(*.xlsx;*.xls)|*.xlsx; *.xls";
    openDlg->Title = "Открыть файл с путями...";

    if (System::Windows::Forms::DialogResult::OK == openDlg->ShowDialog())
    {
        fileName = openDlg->FileName;
        Microsoft::Office::Interop::Excel::Application^ ExcelApplication = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();

        Workbook^ wb = ExcelApplication->Workbooks->Open(fileName, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
        Worksheet^ exWs = safe_cast<Worksheet^>(ExcelApplication->ActiveSheet);

        // Найти последнюю заполненую строку
        int lastUsedRow = exWs->Cells->Find("*", System::Reflection::Missing::Value,
            System::Reflection::Missing::Value, System::Reflection::Missing::Value,
            Microsoft::Office::Interop::Excel::XlSearchOrder::xlByRows, Microsoft::Office::Interop::Excel::XlSearchDirection::xlPrevious,
            false, System::Reflection::Missing::Value, System::Reflection::Missing::Value)->Row;

        vector<string> ArrayTableStrBg;
        vector<string> ArrayTableStrEnd;
        String^ tmp;
        string buff;
        int col = 1;
        for (int i = 1; i <= lastUsedRow; i++) {
            col = 1;
            tmp = ((Microsoft::Office::Interop::Excel::Range^)exWs->Cells[(System::Object^)i, (System::Object^)col])->Value2->ToString();
            buff = msclr::interop::marshal_as<std::string>(tmp);
            ArrayTableStrBg.push_back(buff);
        }
        col = 2;
        for (int i = 1; i <= lastUsedRow; i++) {
            tmp = ((Microsoft::Office::Interop::Excel::Range^)exWs->Cells[(System::Object^)i, (System::Object^)col])->Value2->ToString();
            buff = msclr::interop::marshal_as<std::string>(tmp);
            ArrayTableStrEnd.push_back(buff);
        }
        for (int i = 0; i < lastUsedRow; i++) {
            g.AddWay(ArrayTableStrBg[i], ArrayTableStrEnd[i]);
        }
        ExcelApplication->Workbooks->Close();
    }
    else
        exit(0);
}

[STAThread]
int main()
{
    ExcelAddVertexs();
    ExcelAddWay();
    setlocale(LC_ALL, "RU");
    char key;
    vector<vector<int>> Table = g.CreateGraph();
    do
    {
        system("cls");
        DisplayMenu();
        key = _getch();
        system("cls");
        switch (key)
        {
        case '1': {     //алгоритм поиска в ширину
            string bgName, endName;
            cout << "Введите назвние начальной Вершины: ";
            cin >> bgName;
            cout << "Введите назвние конченой Вершины: ";
            cin >> endName;
            g.BFC(Table, bgName, endName);
            cout << endl;
            break;
        }
        case '2': {     //Добавление вершины
            string nameVertex;
            cout << "Введите назвние новой Вершины: ";
            cin >> nameVertex;
            bool res = g.CreateVertex(nameVertex);
            if (res == true) {
                cout << endl << "Вершина была создана.. " << endl;
                Table = g.CreateGraph();
            }
            break;
        }
        case '3': {      //Добавление ребра
            string bgName, endName;
            int weight;
            cout << "Введите назвние начальной Вершины: ";
            cin >> bgName;
            cout << "Введите назвние конченой Вершины: ";
            cin >> endName;
            bool res = g.AddWay(bgName, endName);
            if (res == true) {
                cout << endl << "Новый путь добавлен..." << endl;
                Table = g.CreateGraph();
            }
            break;
        }
        case '4': {     //Отображение таблицы смежности
            g.DisplayTableGraph(Table);
            break;
        }
        case 27: {
            return 0;
            break;
        }

        default:
            break;
        }

        cout << endl;

        system("pause");
    } while (key != 27);
    return 0;
}