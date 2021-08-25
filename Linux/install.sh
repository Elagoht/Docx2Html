#/usr/bin/bash
sudo cp ./bin/docx2html /usr/bin
sudo mkdir -pv /usr/share/Docx2Html
sudo cp docx2html.png /usr/share/Docx2Html
sudo cp Docx2Html.appimage /usr/share/Docx2Html
sudo cp Docx2Html.desktop /usr/share/applications
echo "Installation has been completed."