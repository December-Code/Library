#include <iostream>
#include <fstream>

using namespace std;
int main()
{
    fstream file;
    int count = 0;
    // string data =fmt::;
    file.open("Web.txt", ios::out);
    if (file.fail())
        cout << "Can't open the file";
    else
    {
        for (int i = 100; i > count; i--)
        {
            file << "<td><a href=";
            file << "\"https://library-r.ntust.edu.tw/var/file/49/1049/img/2763/chiall10905.pdf\"" << endl;
            file << "target=\"_blank\" title=\"0" + to_string(i) + "\">0" + to_string(i) + "期</a></td>" << endl;
        }
        file.close();
    }

    return 0;
}