#pragma once
#include "Header.h"

struct Way
{
    string endVertex;
    int weight;

    Way(string pnEndVertex) {
        endVertex = pnEndVertex;
        weight = 1;
    }
};

class Vertex
{
private:
    string name;
    int count;
    vector<Way> ways;
public:
    Vertex(string pnName) {
        name = pnName;
        count = 0;
    };

    void setName(string pnName) { name = pnName; };
    string getName() { return name; };

    int getCount() { return count; }

    Way getWay(int pnWayID) { return ways[pnWayID]; }

    void AddWay(string pnEndVertex) {
        ways.push_back(Way(pnEndVertex));
        count++;
    }
};

class Graph {
private:
    vector<Vertex> vertexArray;

public:
    vector<vector<int>> CreateGraph() {         // �������� ������� ���������
        vector<vector<int>> Table(vertexArray.size());
        for (int i = 0; i < vertexArray.size(); i++) {
            Table[i].resize(vertexArray.size());
        }
        for (int i = 0; i < vertexArray.size(); i++) {
            for (int j = 0; j < vertexArray.size(); j++) {
                for (int k = 0; k < vertexArray[i].getCount(); k++) {
                    if (vertexArray[j].getName() == vertexArray[i].getWay(k).endVertex)
                        Table[i][j] = 1;
                    else if (Table[i][j] == NULL) {
                        Table[i][j] = 0;
                    }
                }

            }
        }
        return Table;
    }

    void DisplayTableGraph(vector<vector<int>> pnTable) {       // ����������� ������� ���������
        for (int i = 0; i < vertexArray.size(); i++)
        {
            if (i == 0)
            {
                cout << "    " << setw(4) << " " << setw(4) << " |";
            }
            cout << setw(4) << vertexArray[i].getName() << setw(4);
        }
        cout << endl;
        for (int i = 0; i < vertexArray.size(); i++) {
            cout << " " << setw(4) << vertexArray[i].getName() << setw(4) << " |";
            for (int j = 0; j < vertexArray.size(); j++) {

                cout << setw(4) << pnTable[i][j] << setw(4);
            }
            cout << endl;
        }
        cout << endl;
    }

    bool CreateVertex(string pnName) {
        bool checkBoxName = true;
        int i = 0;
        while (checkBoxName && i < vertexArray.size()) {
            if (vertexArray[i].getName() == pnName)
                checkBoxName = false;
            i++;
        }
        if (checkBoxName) {
            vertexArray.push_back(Vertex(pnName));
        }
        else {
            cout << "���������� ������� ������� c ������ " << pnName << endl << "������� � ����� ������ ��� ����������" << endl;
            return false;
        }
        return true;
    }

    bool AddWay(string pnBeginVertex, string pnEndVertex) {
        bool checkBoxBeg = true;
        bool checkBoxEnd = true;
        for (int i = 0; i < vertexArray.size(); i++) {
            if (vertexArray[i].getName() == pnBeginVertex) {
                vertexArray[i].AddWay(pnEndVertex);
                checkBoxBeg = false;
            }
            if (vertexArray[i].getName() == pnEndVertex) {
                checkBoxEnd = false;
            }
        }
        if (checkBoxBeg || checkBoxEnd) {
            cout << "����: \"" << pnBeginVertex << "\" - \"" << pnEndVertex << "\"" << "�� ��� ������!\n";
            cout << "������� �������� �������� ������";
            return false;
        }
        return true;
    }

    void BFC(vector<vector<int>> pnTable, string pnFinalPoint, string pnBeginPoint) {
        int n = pnTable.size(); // ����� ������

        int BeginPoint = -1, FinalPoint = -1;
        for (int i = 0; i < vertexArray.size(); i++) {
            if (vertexArray[i].getName() == pnFinalPoint) {
                BeginPoint = i;
            }
            else if (vertexArray[i].getName() == pnBeginPoint)
                FinalPoint = i;
        }
        if (BeginPoint == -1 && FinalPoint == -1) {
            cout << "��������� � �������� ������� �� �������";
        }
        else if (BeginPoint == -1) {
            cout << "��������� ������� �� �������";
        }
        else if (FinalPoint == -1) {
            cout << "�������� ������� �� �������";
        }
        else {
            queue <int> q; // ������� ��� �������� ������� ������

            vector<bool> visited(pnTable.size()); //false - ������� �� �����������, true - �����������
            vector<bool> inqueue(pnTable.size()); //false - ������� �� � �������, true - � �������
            int start = BeginPoint; // ����� ��������� �������
            int cur; // ����� ���������� �������
            vector <int> prev(vertexArray.size());
            q.push(start); // ���������� ��������� ������� � �������
            visited[start] = inqueue[start] = true;
            while (!q.empty())
            {
                cur = q.front();
                visited[cur] = true;
                q.pop();
                for (int i = 0; i < pnTable.size(); i++)
                {
                    if (!inqueue[i] && pnTable[cur][i])
                    {
                        q.push(i);
                        inqueue[i] = true;
                        prev[i] = cur;
                    }
                }
            }

            vector<int> path;
            for (int v = FinalPoint; v != BeginPoint; v = prev[v])
                path.push_back(v);
            path.push_back(BeginPoint);
            reverse(path.begin(), path.end());

            cout << "����: ";
            for (int i = 0; i < path.size(); i++) {
                if (i == (path.size() - 1))
                    cout << "\"" << vertexArray[path[i]].getName() << "\"";
                else
                    cout << "\"" << vertexArray[path[i]].getName() << "\"" << " -> ";
            }
        }
    }
};
