using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using Word = Microsoft.Office.Interop.Word;

namespace Medicine2.WordData
{
    public static class WordIO
    {
 

        public static List<string> GetDocumentInfo(int Id, string pExtension)
        {
            List<string> lResult = new List<string>();


            if (pExtension.Substring(0,4) != ".doc")
            {
                lResult.Add("Invalid file extension");
                return lResult;
            }


            Word.Application lApp = new Word.Application();
            Word.Document lDocument;

            try
            {
                ExtractDocument(Id);

                //**************************
                List<string> lNappies = new List<string>();
                object ReadOnly = true;
                object ConfirmConversions = true;
                lDocument = lApp.Documents.Open(@"C:/Subs/Leaflet" + pExtension, ref ConfirmConversions, ref ReadOnly);
                lDocument.Activate();

                Word.Paragraphs lParagraphs = lDocument.Paragraphs;

                if (lParagraphs.Count == 0)
                {
                    lResult.Add("No Paragraphs");
                    return lResult;
                }

                Word.Paragraph lParagraph = lDocument.Paragraphs.First;
                Word.Range lRange = lParagraph.Range;
                lRange.Select();
                lResult.Add(lApp.Selection.Text);

                // Search for Nappi codes 

                if (lParagraphs.Count > 3)
                {
                    Regex Test1 = new Regex(@"^[\d]{*}-[\d]{3}[\w]*");
                   
                    foreach (Word.Paragraph item in lDocument.Paragraphs)
                    {
                        Word.Range lRange2 = item.Range;
                        lRange.Select();
                        string lTarget = lApp.Selection.Text.Substring(0, 60);
  
                        if (Test1.IsMatch(lTarget))
                        {
                            lResult.Add(lTarget);
                        }
                    }
                }

                return lResult;


                //***********************

                //{
                //    //Console.WriteLine(lTarget);

                //    string lTarget2 = lTarget.Substring(lTarget.IndexOf('R') + 1);
                //    lTarget2 = Regex.Replace(lTarget2, @"\s+", String.Empty);
                //    lTarget2 = Regex.Replace(lTarget2, @"\r+", String.Empty);
                //    //lTarget2 = lTarget2.Replace(',', '.');
                //    decimal Price = Decimal.Parse(lTarget2);

                //    //Regex Test2 = new Regex(@"[\d]*[\s]*[\d]+,[\d]{2}");
                //    //if (!Test2.IsMatch(lTarget2))
                //    //{
                //    //    Console.WriteLine("No match for price");
                //    //}
                //    //else
                //    //{ 

                //    //Match lMatch = Test2.Match(lTarget, 7);

                //    Console.WriteLine(gApp.Selection.Text.Substring(0, 10) + "\t" + Price.ToString("#####0.00"));

                //}

                //}


                bool ExtractDocument(int WordId)
                {
                    try
                    {

                        string lQuery = "select Content from WordDocument where Id = " + WordId.ToString();
                        FileStream lWriter = new FileStream(@"c:\Subs\Leaflet" + pExtension, FileMode.Create, FileAccess.ReadWrite);

                        using (SqlConnection lConnection = new SqlConnection( @"Data Source = pklwebdb01\mssql2016std; Initial Catalog = Medicine; Integrated Security = True"))
                        {
                            SqlCommand lCommand = new SqlCommand(lQuery, lConnection);
                            lCommand.CommandType = CommandType.Text;
                            lConnection.Open();
                            SqlDataReader lReader = lCommand.ExecuteReader();
                            lReader.Read();
                            var lResult = lReader.GetValue(0);
                            byte[] lBytes = (byte[])lResult;
                            lWriter.Write(lBytes, 0, lBytes.Length);

                            lWriter.Flush();
                            lWriter.Close();
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        //Display all the exceptions

                        Exception CurrentException = ex;
                        int ExceptionLevel = 0;
                        do
                        {
                            ExceptionLevel++;
                            ExceptionData.WriteException(1, ExceptionLevel.ToString() + " " + CurrentException.Message, "static WordIO", "ExtractDocument", Id.ToString());
                            CurrentException = CurrentException.InnerException;
                        } while (CurrentException != null);

                        throw ex;
                    }
                }

            }

           
            catch (Exception ex)
            {
                //Display all the exceptions

                Exception CurrentException = ex;
                int ExceptionLevel = 0;
                do
                {
                    ExceptionLevel++;
                    ExceptionData.WriteException(1, ExceptionLevel.ToString() + " " + CurrentException.Message,"static Word.IO", "GetFirstParagraph", "");
                    CurrentException = CurrentException.InnerException;
                } while (CurrentException != null);

                throw ex;
            }

            finally
            {
                lApp.Quit();
            }

        }

        //public void Replace()
        //{
        //    gDocument = gApp.Documents.Open(@"D:/Word/Medicine.docx");
        //    gDocument.Activate();

        //    Word.Find findObject = gApp.Selection.Find;
        //    findObject.ClearFormatting();
        //    findObject.Text = "find me";
        //    findObject.Replacement.ClearFormatting();
        //    findObject.Replacement.Text = "Found";

        //    object replaceAll = Word.WdReplace.wdReplaceAll;
        //    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
        //        ref missing, ref missing, ref missing, ref missing, ref missing,
        //        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        //  }



        //public void Show()
        //{
        //    gDocument = gApp.Documents.Open(@"D:/Word/Medicine.docx");
        //    gDocument.Activate();



        //    try
        //    {
        //        //Console.WriteLine(gDocument.Paragraphs.Count.ToString());

        //        Word.Paragraph lParagraph = gDocument.Paragraphs[7];

        //        //object start = 0;
        //        //object end = 7;
        //        Word.Range rng = lParagraph.Range;

        //        rng.Select();
              
        //        //object isVisible = false;
    
             
                
        //        Console.WriteLine(gApp.Selection.Words.Count.ToString());
        //        Console.ReadLine();

        //        //foreach (var item in gApp.Selection.Words)
        //        //{
        //        //    Console.WriteLine(item.ToString());
        //        //}
        //        //Console.ReadLine();

        //        object findText = "R*";
        //        object WildCards = true;



        //        gApp.Selection.Find.ClearFormatting();

        //        if (gApp.Selection.Find.Execute(ref findText,
        //            ref missing, ref missing, ref WildCards, ref missing, ref missing, ref missing,
        //            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
        //            ref missing, ref missing))
        //        {
        //            Console.WriteLine("Text found.");
        //            Console.WriteLine(gApp.Selection.Find.ToString());
        //        }
        //        else
        //        {
        //            Console.WriteLine("The text could not be located.");
        //        }

        //        Console.ReadLine();
        //    }

        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }
        //    finally
        //    {
        //        gDocument.Close();
        //    }

        //}

    } }
