# Pre-Requisites:
1. Visual Studio 2022
2. .NET Core 6 SDK
3. Asponse Word License

# Compiling Program
1. Open *.sln file in Visual studio
2. Press Ctrl + F5 to compile & run application.
3. Navigate to bin\Debug directory and execute the below commands 

## Finding the coordinates of {{Token}}:
1. Open CMD and navigate to the directory where app is placed.
2. Execute command
DocLocationFinder-netcore wordxy "c:\documents\sample.docx"
3. Output coordinates will be shown for each occurance on console screen.

## Removing {{Token}} from word document:
1. Open CMD and navigate to the directory where app is placed.
2. Execute command
DocLocationFinder-netcore wordremove "c:\documents\sample.docx"
3. Output coordinates will be shown for each occurance on console screen.

## Convert Word document to PDF:
1. Open CMD and navigate to the directory where app is placed.
2. Execute command
DocLocationFinder-netcore wordtopdf "c:\documents\sample.docx"
3. Output coordinates will be shown for each occurance on console screen.