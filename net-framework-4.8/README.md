# Pre-Requisites:
1. Visual Studio 2022
2. .NET Framework 4.8 Developer Pack
3. Asponse Word License

# Compiling Program
1. Open *.sln file in Visual studio
2. Press Ctrl + F5 to compile & run application.
3. Navigate to bin\Debug directory and execute the below commands 

## Finding the coordinates of {{Token}}:
1. Open CMD and navigate to the directory where app is placed.
2. Execute command
DocLocationFinder wordxy "c:\documents\sample.docx"
3. Output coordinates will be shown for each occurance on console screen.

## Removing {{Token}} from word document:
1. Open CMD and navigate to the directory where app is placed.
2. Execute command
DocLocationFinder wordremove "c:\documents\sample.docx"
3. Output coordinates will be shown for each occurance on console screen.

## Convert Word document to PDF:
1. Open CMD and navigate to the directory where app is placed.
2. Execute command
DocLocationFinder wordtopdf "c:\documents\sample.docx"
3. Output coordinates will be shown for each occurance on console screen.

## Merge PDF Files:
1. Open CMD and navigate to the directory where app is placed.
2. Execute command
DocLocationFinder pdfmerge "c:\documents\master.docx" "c:\documents\child1.docx" "c:\documents\child2.docx"
3. Output coordinates will be shown for each occurance on console screen.