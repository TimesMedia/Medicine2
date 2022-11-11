using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Medicine2.Data;
using Medicine2.WordData;

namespace Medicine2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonImportWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Data.DataSet1.DocumentDataTable lPRAtable = new DataSet1.DocumentDataTable();
                Data.DataSet1TableAdapters.DocumentTableAdapter lPRAAdapter = new Data.DataSet1TableAdapters.DocumentTableAdapter();

                Data.DataSet1.WordDocumentDataTable lWordTable = new DataSet1.WordDocumentDataTable();
                Data.DataSet1TableAdapters.WordDocumentTableAdapter lWordAdapter = new Data.DataSet1TableAdapters.WordDocumentTableAdapter();

                lPRAAdapter.Fill(lPRAtable);   

                MessageBox.Show("Data obtained from PKLMEDMIMS01");

                foreach (Data.DataSet1.DocumentRow item in lPRAtable)
                {
                    Data.DataSet1.WordDocumentRow lNewRow = lWordTable.NewWordDocumentRow();
                    lNewRow.PRAID = item.ID;
                    lNewRow.Format = item.Format;
                    lNewRow.Content= item.Content;
                    lWordTable.Rows.Add(lNewRow);
                }
                lWordAdapter.Update(lWordTable);
                lWordTable.AcceptChanges();
                MessageBox.Show("Loaded " + lWordTable.Rows.Count.ToString() + " rows");
            }

            catch (Exception ex)
            {
                //Display all the exceptions

                Exception? CurrentException = ex;
                int ExceptionLevel = 0;
                do
                {
                    ExceptionLevel++;
                    WordData.ExceptionData.WriteException(1, ExceptionLevel.ToString() + " " + CurrentException.Message, this.ToString(), "Update", "");
                    CurrentException = CurrentException?.InnerException;
                } while (CurrentException != null);
                
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonTryReadWord_Click(object sender, RoutedEventArgs e)
        {
            int lCurrentRecord = 0;
            this.Cursor = Cursors.Wait;

            try 
            { 
                Medicine2.Data.DataSet1.WordDocumentDataTable lWordTable = new Medicine2.Data.DataSet1.WordDocumentDataTable();
                Data.DataSet1TableAdapters.WordDocumentTableAdapter lWordAdapter = new Data.DataSet1TableAdapters.WordDocumentTableAdapter();
                lWordAdapter.FillBy(lWordTable);

                foreach (DataSet1.WordDocumentRow  lRow in lWordTable.Rows)
                {
                    string? lExtension = lRow.Format;
                    if (!lExtension.Contains(".doc") )
                    {
                        continue;
                    }
                            
                    lCurrentRecord = lRow.Id;
                    string lFirstParagraph = WordIO.GetFirstParagraph(lCurrentRecord, lExtension);
                    if (lFirstParagraph.Length < 2)
                    {
                        lFirstParagraph = "Nothing in first paragraph";
                    }

                    lRow.InitialText = lFirstParagraph;
                    lWordAdapter.Update(lRow); 
                   
                    //MessageBox.Show(lCurrentRecord.ToString() + " " + lFirstParagraph);
                }
                MessageBox.Show(lWordTable.Rows.Count.ToString() + " records updated.");


            }
            catch (Exception ex)
            {
                //Display all the exceptions

                Exception? CurrentException = ex;
                int ExceptionLevel = 0;
                do
                {
                    ExceptionLevel++;
                    WordData.ExceptionData.WriteException(1, ExceptionLevel.ToString() + " " + CurrentException.Message, this.ToString(), "buttonTryReadWord_Click", "CurrentRecord = " + lCurrentRecord.ToString());
                    CurrentException = CurrentException?.InnerException;
                } while (CurrentException != null);

                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Arrow;
            }

        }
    }
}
