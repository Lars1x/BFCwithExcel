#include "pch.h"
#include "Classes.h"
#include "Header.h"

using namespace Microsoft::Office::Interop::Excel;
using namespace System::Runtime::InteropServices;

Graph g;


void DisplayMenu() {
    cout << "\t ����" << endl;
    cout << "1 - �������� ������ � ������" << endl;
    cout << "2 - �������� ����� �������" << endl;
    cout << "3 - �������� ����� ����" << endl;
    cout << "4 - ���������� ������� ���������" << endl;
    cout << endl;
    cout << "Esc - ����� �� ���������" << endl;
}

void ExcelAddVertexs() {
    String^ fileName;
    OpenFileDialog^ openDlg = gcnew OpenFileDialog();

    openDlg->Multiselect = false;
    openDlg->Filter = "Excel Document(*.xlsx;*.xls)|*.xlsx; *.xls";
    openDlg->Title = "������� ���� � ���������...";

    if (System::Windows::Forms::DialogResult::OK == openDlg->ShowDialog())
    {
        fileName = openDlg->FileName;

        Microsoft::Office::Interop::Excel::Application^ ExcelApplication = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();

        Workbook^ wb = ExcelApplication->Workbooks->Open(fileName, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
        Worksheet^ exWs = safe_cast<Worksheet^>(ExcelApplication->ActiveSheet);

        // ����� ��������� ���������� ������
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
    openDlg->Title = "������� ���� � ������...";

    if (System::Windows::Forms::DialogResult::OK == openDlg->ShowDialog())
    {
        fileName = openDlg->FileName;
        Microsoft::Office::Interop::Excel::Application^ ExcelApplication = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();

        Workbook^ wb = ExcelApplication->Workbooks->Open(fileName, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
        Worksheet^ exWs = safe_cast<Worksheet^>(ExcelApplication->ActiveSheet);

        // ����� ��������� ���������� ������
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
        case '1': {     //�������� ������ � ������
            string bgName, endName;
            cout << "������� ������� ��������� �������: ";
            cin >> bgName;
            cout << "������� ������� �������� �������: ";
            cin >> endName;
            g.BFC(Table, bgName, endName);
            cout << endl;
            break;
        }
        case '2': {     //���������� �������
            string nameVertex;
            cout << "������� ������� ����� �������: ";
            cin >> nameVertex;
            bool res = g.CreateVertex(nameVertex);
            if (res == true) {
                cout << endl << "������� ���� �������.. " << endl;
                Table = g.CreateGraph();
            }
            break;
        }
        case '3': {      //���������� �����
            string bgName, endName;
            int weight;
            cout << "������� ������� ��������� �������: ";
            cin >> bgName;
            cout << "������� ������� �������� �������: ";
            cin >> endName;
            bool res = g.AddWay(bgName, endName);
            if (res == true) {
                cout << endl << "����� ���� ��������..." << endl;
                Table = g.CreateGraph();
            }
            break;
        }
        case '4': {     //����������� ������� ���������
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