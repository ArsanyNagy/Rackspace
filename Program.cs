using Microsoft.Office.Interop.Excel;
using RackspaceTask.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelImporter;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace RackspaceTask
{
    class Program
    {
        public static List<Book> bookList = new List<Book>();
        static void Main(string[] args)
        {
            //starting functions 
            //first: read excel file
            //second: show to user choice options
            ReadExcel();
            MainChooseOption();
        }

        private static void ReadExcel()
        {
            //open app
            try
            {
                Application excelApp = new Application();
                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not installed!!");
                    return;
                }
                //open excel sheet
                var path = Path.Combine(Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()) + "\\Task.xlsx");
                Workbook excelBook = excelApp.Workbooks.Open(@"" + path + "");
                _Worksheet excelSheet = excelBook.Sheets[1];
                Range excelRange = excelSheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                
                //append data from excel sheet to global variable
                bookList = new List<Book>();

                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    Book book = new Book();
                    book.id = int.Parse(excelRange.Cells[i, 1]?.Value2.ToString());
                    book.title = excelRange.Cells[i, 2].Value2?.ToString();
                    book.author = excelRange.Cells[i, 3].Value2?.ToString();
                    book.description = excelRange.Cells[i, 4].Value2?.ToString();
                    bookList.Add(book);
                }
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("something wrong happened press any key to close app");
                Console.ReadLine();
                SaveAndExit();
            }
        }
        private static void MainChooseOption()
        {
            //choices for user
            Console.WriteLine("==== Book Manager ==== \r");
            Console.WriteLine("1) View All books ");
            Console.WriteLine("2) Add a book ");
            Console.WriteLine("3) Edit a book ");
            Console.WriteLine("4) Search for a book ");
            Console.WriteLine("5) Save and exit ");
            Console.Write("Choose [1-5]: ");
            var choose = Console.ReadLine();
            var valueChoosen = 0;
            if (int.TryParse(choose, out valueChoosen) && valueChoosen < 6 && valueChoosen > 0)
            {
                if (valueChoosen == 1)
                    ViewBooks();
                else if (valueChoosen == 2)
                    AddBook();
                else if (valueChoosen == 3)
                    EditBook();
                else if (valueChoosen == 4)
                    SearchBook();
                else if (valueChoosen == 5)
                    SaveAndExit();
            }
        }
        private static void ViewBooks()
        {
            //for loop on glocal variable to show books
            foreach (var book in bookList)
            {
                Console.WriteLine("[" + book.id + "] " + book.title + "");
            }
            //to show specific book
            ShowBook(bookList);
        }
        private static void AddBook()
        {
            //for add book
            var bookData = new Book();
            Console.WriteLine("==== Add a Book ==== ");
            Console.WriteLine("Please enter the following information: ");
            Console.Write("Title : ");
            var bookTitle = Console.ReadLine();
            bookData.title = bookTitle;
            Console.Write("Author : ");
            var bookAuthor = Console.ReadLine();
            bookData.author = bookAuthor;
            Console.Write("Description : ");
            var bookDesc = Console.ReadLine();
            bookData.description = bookDesc;
            //if title and author fields not empty add book
            //else choose another option
            if (bookTitle != string.Empty && bookAuthor != string.Empty)
                updateExcel(bookData, "Add");
            else
                MainChooseOption();
        }
        private static void EditBook()
        {
            //for loop on glocal variable to show books
            Console.WriteLine("==== Edit a Book ==== ");
            foreach (var book in bookList)
            {
                Console.WriteLine("[" + book.id + "] : " + book.title + "");
            }
            //for show specific book for edit
            ShowBookForEdit();
        }
        private static void SearchBook()
        {
            //search book by title 
            Console.WriteLine("==== Search ==== ");
            Console.WriteLine("Type in one or more keywords to search for ");
            Console.Write("Search : ");
            var searchValue = Console.ReadLine();
            Console.WriteLine("The following books matched your query. Enter the book ID to see more details, or <Enter> to return. ");
            var booksData = bookList.Where(x => x.title.ToLower().Contains(searchValue.ToLower())).ToList();
            //for loop on all books in search result
            if (booksData.Count != 0)
            {
                foreach (var book in booksData)
                {
                    Console.WriteLine("[" + book.id + "] : " + book.title + "");
                }
                //for show specific book
                ShowBook(booksData);
            }

        }
        private static void SaveAndExit()
        {
            //close app
            Environment.Exit(0);
        }
        private static void updateExcel(Book bookData, string isAddOrEdit)
        {
            //open excel file
            Application excelApp = new Application();
            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }
            var path = Path.Combine(Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()) + "\\Task.xlsx");
            Workbook excelBook = excelApp.Workbooks.Open(@"" + path + "");
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            //if Edit mode update specific book
            //else Add mode add book
            if (isAddOrEdit == "Edit")
            {
                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    int cellValue = int.Parse(excelRange.Cells[i, 1]?.Value2.ToString());
                    if (cellValue == bookData.id)
                    {
                        excelRange.Cells.set_Item(i, 2, bookData.title);
                        excelRange.Cells.set_Item(i, 3, bookData.author);
                        excelRange.Cells.set_Item(i, 4, bookData.description);
                    }
                }
            }
            else if (isAddOrEdit == "Add")
            {
                bookData.id = int.Parse(excelRange.Cells[rowCount, 1].Value2.ToString()) + 1;
                excelRange.Cells.set_Item(rowCount + 1, 1, bookData.id);
                excelRange.Cells.set_Item(rowCount + 1, 2, bookData.title);
                excelRange.Cells.set_Item(rowCount + 1, 3, bookData.author);
                excelRange.Cells.set_Item(rowCount + 1, 4, bookData.description);
            }
            //save data
            excelBook.Save();
            excelApp.Quit();
            Console.WriteLine("Book [" + bookData.id + "] Saved.");
            //to update global variable
            ReadExcel();
            //if edit mode show book
            //else show choices 
            if (isAddOrEdit == "Edit")
            {
                ShowBookForEdit();
            }
            else if (isAddOrEdit == "Add")
            {
                MainChooseOption();
            }
        }
        private static void ShowBook(List<Book> booksData)
        {
            //show book by id
            Console.WriteLine("To view details enter the book ID, to return press <Enter>.  ");
            Console.Write("Book ID: ");
            var choose = Console.ReadLine();
            var bookChoosen = 0;
            if (int.TryParse(choose, out bookChoosen) && booksData.Any(x => x.id == bookChoosen))
            {
                var bookValue = booksData.Where(x => x.id == bookChoosen).FirstOrDefault();
                if (bookValue != null)
                {
                    Console.WriteLine("ID : " + bookValue.id + " ");
                    Console.WriteLine("Title : " + bookValue.title + " ");
                    Console.WriteLine("Author : " + bookValue.author + " ");
                    Console.WriteLine("Description : " + bookValue.description + " ");
                }
                ShowBook(booksData);
            }
            else
            {
                //if book id wrong show choices
                MainChooseOption();
            }
        }
        private static void ShowBookForEdit()
        {
            //show book by id
            Console.WriteLine("Enter the book ID of the book you want to edit; to return press <Enter>.  ");
            Console.Write("Book ID: ");
            var choose = Console.ReadLine();
            var bookChoosen = 0;
            if (int.TryParse(choose, out bookChoosen) && bookList.Any(x => x.id == bookChoosen))
            {
                Console.WriteLine("Input the following information. To leave a field unchanged, hit <Enter> ");

                var bookValue = bookList.Where(x => x.id == bookChoosen).FirstOrDefault();
                Console.Write("Title [" + bookValue.title + "]: ");
                var bookTitle = Console.ReadLine();
                bookValue.title = bookTitle == string.Empty ? bookValue.title : bookTitle;
                Console.Write("Author [" + bookValue.author + "]: ");
                var bookAuthor = Console.ReadLine();
                bookValue.author = bookAuthor == string.Empty ? bookValue.author : bookAuthor;
                Console.Write("Description [" + bookValue.description + "]: ");
                var bookDesc = Console.ReadLine();
                bookValue.description = bookDesc == string.Empty ? bookValue.description : bookDesc;
                updateExcel(bookValue, "Edit");
                ShowBookForEdit();
            }
            else
            {
                //if book id wrong show choices
                MainChooseOption();
            }
        }
    }
}
