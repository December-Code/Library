#include <iostream>
#include <fstream>
#include <iomanip>

using namespace std;
int main()
{
    fstream file;
    int countS; //從哪開始
    int countE; //到哪結束
    cout << "Input which number start?" << endl;
    cin >> countS;
    cout << "Input which number End?" << endl;
    cin >> countE;

    file.open("Web.txt", ios::out);
    if (file.fail())
        cout << "Can't open the file";
    else
    {
        int count_add = 0;
        for (int i = countE; i > countS - 1; i--)
        {
            if (count_add % 4 == 0)
            {
                file << "<tr>" << endl;
            }
            file << "<td><a href=";
            file << "\"https://library-r.ntust.edu.tw/var/file/49/1049/img/2781/Newsletter";
            file << setw(3) << setfill('0') << to_string(i);
            file << ".pdf\"" << endl;
            file << "target=\"_blank\" title=\"";
            file << setw(3) << setfill('0') << to_string(i);
            file << "\">";
            file << setw(3) << setfill('0') << to_string(i);
            file << "期</a></td>" << endl;
            if (count_add % 4 == 3 or i == countS)
            {
                file << "</tr>" << endl;
            }
            count_add += 1;
        }
        file.close();
        cout << "Completed" << endl;
    }
    system("pause");
    return 0;
}