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
    int count_add = 0;

    file.open("Web.txt", ios::out);
    if (file.fail())
        cout << "Can't open the file";
    else
    {
        // for (int i = countE; i > countS - 1; i--)
        // {
        //     if (count_add % 5 == 0)
        //     {
        //         file << "<li><a href=\"#\">" + to_string(i) + "~" + to_string(i - 4) + "期</a>" << endl;
        //         file << "<ul>" << endl;
        //     }
        //     file << "<li><a href=";
        //     file << "\"https://library-r.ntust.edu.tw/var/file/49/1049/img/2781/Newsletter";
        //     file << setw(3) << setfill('0') << to_string(i);
        //     file << ".pdf\"" << endl;
        //     file << "target=\"_blank\" title=\"";
        //     file << setw(3) << setfill('0') << to_string(i);
        //     file << "\">";
        //     file << setw(3) << setfill('0') << to_string(i);
        //     file << "期</a></li>" << endl;
        //     if (count_add % 5 == 4 or i == countS)
        //     {
        //         file << "</ul>" << endl;
        //         file << "</li>" << endl;
        //     }
        //     count_add += 1;
        // }
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