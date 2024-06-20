﻿using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Replacing;
using DocLocationFinder.Common;
using DocLocationFinder.Helpers;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace DocLocationFinder
{
    class Program
    {
        static void Main(string[] args)
        {
            ApplyLicense();

            var isValid = ValidateInput(args);
            
            if (!isValid)
            {
                return;
            }

            var path = args[1];


            if (args[0] == AppConstants.Action.FindWordFileCoordinates)
            {
                Console.WriteLine("Command Action: Find Coordinates XY for Word document.");
                Console.WriteLine();
            }
            else if (args[0] == AppConstants.Action.ReplaceWordFileText)
            {
                Console.WriteLine("Command Action: Removing {{Token}} from Word document.");
                Console.WriteLine();
            }
            else if (args[0] == AppConstants.Action.ConvertWordToPdf)
            {
                Console.WriteLine("Command Action: Converting Word document into PDF document.");
                Console.WriteLine();
            }


            Console.WriteLine();
            if (args[0] == AppConstants.Action.FindWordFileCoordinates)
            {


                if (args.Length < 2)
                {
                    Console.WriteLine("Please provide the required commands to proceed.");
                    Console.WriteLine("DocLocationFinder [param1] [param2]");
                    Console.WriteLine("[param1] should be the action. (i.e 'wordxy')");
                    Console.WriteLine("[param2] should be the file path. (i.e 'C:\\documents\\sample.docx')");
                    Console.WriteLine("");
                    isValid = false;
                }

                if (!isValid)
                {
                    Console.WriteLine("There were some validation errors so the process could not be started.");
                    Console.ReadLine();
                    return;
                }

                Document doc = new Document(path);

                //Find the text between <<>> and insert bookmark
                doc.Range.Replace(new Regex(@"\{\{(.*?)\}\}"), "", new FindReplaceOptions() { ReplacingCallback = new FindAndInsertBookmark() });

                LayoutCollector layoutCollector = new LayoutCollector(doc);
                LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
                //Display the left top position of text between angle bracket.
                int bookmarkCount = 0;
                DocumentBuilder builder = new DocumentBuilder(doc);
                foreach (Bookmark bookmark in doc.Range.Bookmarks)
                {
                    if (bookmark.Name.StartsWith("bookmark_"))
                    {
                        bookmarkCount++;
                        layoutEnumerator.Current = layoutCollector.GetEntity(bookmark.BookmarkStart);
                        int pageNo = layoutCollector.GetStartPageIndex(bookmark.BookmarkStart);

                        Paragraph paragraph = bookmark.BookmarkStart.ParentNode as Paragraph;
                        string paragraphText = paragraph.GetText();
                        Console.WriteLine($"Page No: {pageNo}, X= {layoutEnumerator.Rectangle.Left}, Y= {layoutEnumerator.Rectangle.Top}    =>    Text= '{paragraphText.Trim()}'");

                    }
                }



                Console.WriteLine("------------------------------------");
                Console.WriteLine("No. of tokens found: " + bookmarkCount);
                Console.WriteLine("------------------------------------");
            }

            else if (args[0] == AppConstants.Action.ReplaceWordFileText)
            {

                Document doc = new Document(path);

                if (args.Length < 2)
                {
                    Console.WriteLine("Please provide the required commands to proceed.");
                    Console.WriteLine("DocLocationFinder [param1] [param2]");
                    Console.WriteLine("[param1] - Action. (i.e wordremove)");
                    Console.WriteLine("[param2] - Word file path. (i.e 'C:\\documents\\sample.docx')");
                    Console.WriteLine("");
                    isValid = false;
                }

                if (!isValid)
                {
                    Console.WriteLine("There were some validation errors so the process could not be started.");
                    Console.ReadLine();
                    return;
                }

                //Find the text between <<>> and insert bookmark
                doc.Range.Replace(new Regex(@"\{\{(.*?)\}\}"), "", new FindReplaceOptions() { ReplacingCallback = new ReplaceAndInsertBookmark() });

                LayoutCollector layoutCollector = new LayoutCollector(doc);
                LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
                //Display the left top position of text between angle bracket.
                int bookmarkCount = 0;
                DocumentBuilder builder = new DocumentBuilder(doc);
                foreach (Bookmark bookmark in doc.Range.Bookmarks)
                {
                    if (bookmark.Name.StartsWith("bookmark_"))
                    {
                        bookmarkCount++;
                        layoutEnumerator.Current = layoutCollector.GetEntity(bookmark.BookmarkStart);
                        int pageNo = layoutCollector.GetStartPageIndex(bookmark.BookmarkStart);

                        Paragraph paragraph = bookmark.BookmarkStart.ParentNode as Paragraph;
                        string paragraphText = paragraph.GetText();
                        Console.WriteLine($"Page No: {pageNo}, X= {layoutEnumerator.Rectangle.Left}, Y= {layoutEnumerator.Rectangle.Top}    =>    Text removed. {paragraphText}");

                        // Insert two shapes along with a group shape with another shape inside it.
                        //builder.InsertShape(ShapeType.Rectangle, 60, 60);
                        //builder.InsertShape(ShapeType.Rectangle, 60, 60);

                        //Shape shape = new Shape(doc, ShapeType.Rectangle);

                        //shape.RelativeHorizontalPosition = Aspose.Words.Drawing.RelativeHorizontalPosition.Page;
                        //shape.RelativeVerticalPosition = Aspose.Words.Drawing.RelativeVerticalPosition.Page;
                        //shape.Width = 80;
                        //shape.Height = 15;
                        //shape.FillColor = Color.Transparent;
                        //shape.Left = layoutEnumerator.Rectangle.Left;
                        //shape.Top = layoutEnumerator.Rectangle.Top;
                        builder.MoveToBookmark(bookmark.Name);
                        //builder.InsertNode(shape);
                    }
                }



                Console.WriteLine("------------------------------------");
                Console.WriteLine("Removed occurences: " + bookmarkCount);
                Console.WriteLine("------------------------------------");

                var nameWithoutExt = Path.GetFileNameWithoutExtension(path);
                var extension = Path.GetExtension(path);
                var directory = Path.GetDirectoryName(path);

                var outputDocxFilePath = Path.Combine(directory, nameWithoutExt + "_output" + extension);

                doc.Save(outputDocxFilePath);
                Console.WriteLine($"Output Word File saved at '{outputDocxFilePath}'");
                //System.Diagnostics.Process.Start(outputFilePath);

            }
            else if (args[0] == AppConstants.Action.ConvertWordToPdf)
            {

                Document doc = new Document(path);

                if (args.Length < 2)
                {
                    Console.WriteLine("Please provide the required commands to proceed.");
                    Console.WriteLine("DocLocationFinder [param1] [param2]");
                    Console.WriteLine("[param1] - Action. (i.e wordtopdf)");
                    Console.WriteLine("[param2] - Word file path. (i.e 'C:\\documents\\sample.docx')");
                    Console.WriteLine("");
                    isValid = false;
                }

                if (!isValid)
                {
                    Console.WriteLine("There were some validation errors so the process could not be started.");
                    Console.ReadLine();
                    return;
                }

                var nameWithoutExt = Path.GetFileNameWithoutExtension(path);
                var directory = Path.GetDirectoryName(path);

                var outputPdfFilePath = Path.Combine(directory, nameWithoutExt + "_output" + ".pdf");
                doc.Save(outputPdfFilePath, SaveFormat.Pdf);
                Console.WriteLine($"Output PDF File saved at '{outputPdfFilePath}'");
            }
            Console.ReadLine();
        }

        private static bool ValidateInput(string[] args)
        {
            var isValid = true;
            if (args.Length > 1)
            {
                if (string.IsNullOrEmpty(args[0]))
                {
                    Console.WriteLine("Please define the action. (i.e 'wordxy, 'wordremove', wordtopdf')");
                    isValid = false;
                }
                if (string.IsNullOrEmpty(args[1]))
                {
                    Console.WriteLine("Please provide the file path in 2nd param.");
                    isValid = false;
                }
                if (!string.IsNullOrEmpty(args[1]))
                {
                    var path = args[1];
                    try
                    {
                        if (!File.Exists(path))
                        {
                            Console.WriteLine("File does not exists in provided path.");
                            isValid = false;
                        }
                    }
                    catch
                    {
                        Console.WriteLine("The provided path does not look valid.");
                        isValid = false;
                    }
                }

                if (args[0] != AppConstants.Action.FindWordFileCoordinates
                    && args[0] != AppConstants.Action.ReplaceWordFileText
                    && args[0] != AppConstants.Action.ConvertWordToPdf)
                {
                    Console.WriteLine("The provided action does not look valid.");
                    isValid = false;
                }
            }
            else
            {
                Console.WriteLine("Please provide the required commands to proceed.");
                Console.WriteLine("DocLocationFinder [param1] [param2]");
                Console.WriteLine("[param1] should be the action. (i.e 'wordxy, wordremove, wordtopdf')");
                Console.WriteLine("[param2] should be the file path. (i.e 'C:\\documents\\sample.docx')");
                Console.WriteLine("");
                isValid = false;
            }

            if (!isValid)
            {
                Console.WriteLine("There were some validation errors so the process could not be started.");
                return false;
            }
            return isValid;
        }

        private static void ApplyLicense()
        {
            Aspose.Words.License lic = new Aspose.Words.License();
            lic.SetLicense(@"Aspose.Words.lic");
        }
    }


}
