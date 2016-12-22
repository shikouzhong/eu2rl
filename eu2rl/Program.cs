using System;
using System.Collections.Generic;
using System.Linq;
using MathNet.Numerics;
using MathNet.Numerics.LinearAlgebra;
using System.IO;
using Origin;

namespace eu2rl
{
    class InputFile
    {
        public string FileName;
        public Vector<Complex32> frequency, epsilon, mu; // Frequency in GHz.

        public void ReadCSV()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("#1/3 Pelese select the filename or path of the target file to start...");
            Console.ResetColor();

            string myFilePath = Console.ReadLine();
            if (myFilePath.Contains('.'))
            {
                char[] separatorArray = { '\\', '.', '/' };
                var tempArray = myFilePath.Split(separatorArray);
                FileName = tempArray[tempArray.Length - 2];
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR!!! Must attach the file extension.");
                Console.ResetColor();
                ReadCSV();
            }

            List<Complex32> listFreq = new List<Complex32>();
            List<Complex32> listEpsilon = new List<Complex32>();
            List<Complex32> listMu = new List<Complex32>();

            try
            {
                var sr = new StreamReader(myFilePath);
                if (sr.ReadLine().Contains("Label"))
                {
                    sr.ReadLine();
                }
                while (!sr.EndOfStream)
                {
                    try
                    {
                        char[] sChar = { ',', '\t' };
                        var values = sr.ReadLine().Split(sChar);
                        listFreq.Add(float.Parse(values[0]) / 1.0e9f);
                        listEpsilon.Add(new Complex32(float.Parse(values[1]), -float.Parse(values[2])));
                        listMu.Add(new Complex32(float.Parse(values[3]), -float.Parse(values[4])));
                    }
                    catch (Exception)
                    {
                        // Do nothing here to skip the headers.
                    }
                }
                sr.Close();

                var Vc = Vector<Complex32>.Build;
                this.frequency = Vc.Dense(listFreq.ToArray());
                this.epsilon = Vc.Dense(listEpsilon.ToArray());
                this.mu = Vc.Dense(listMu.ToArray());
            }
            catch (Exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR!!! Files are NOT found.");
                Console.ResetColor();
                this.ReadCSV();
            }
        }

        public Vector<double> CalcRL(float inputThickness)
        {
            //**********************************************************************************************//
            //* z = sqrt(mu./ epsilon).* tanh((j * 2 * pi / c) * thickness .* freq .* sqrt(mu.* epsilon)); *//
            //* rl = 20 * log10(abs((z - 1)./ (z + 1)));                                                   *//
            //**********************************************************************************************//
            // Frequency in GHz, inputThickness in mm. 300 corresponds to the speed of light.
            var sqrt_uDivideE = (mu.PointwiseDivide(epsilon)).Map(x => Complex32.Sqrt(x));
            var sqrt_uMultiplyE = (mu.PointwiseMultiply(epsilon)).Map(x => Complex32.Sqrt(x));
            var thickFactor = (2 * (float)Math.PI * inputThickness / 300) * Complex32.ImaginaryOne;
            var tanh_Factor = (thickFactor * frequency.PointwiseMultiply(sqrt_uMultiplyE)).Map(x => Complex32.Tanh(x));

            var z = sqrt_uDivideE.PointwiseMultiply(tanh_Factor);
            var RL = z.Map(x => (20 * Math.Log10(Complex32.Abs((x - 1) / (x + 1)))));
            return RL;
        }

        public Matrix<double> CalcRL(Vector<float> inputThickness)
        {
            var fCount = frequency.Count;
            var tCount = inputThickness.Count;
            Matrix<double> RL = Matrix<double>.Build.Dense(fCount, tCount);
            for (int i = 0; i < tCount; i++)
            {
                var newColumn = this.CalcRL(inputThickness[i]);
                RL.SetColumn(i, newColumn);
            }
            return RL;
        }
    }


    class OutputFile
    {
        public ApplicationCOMSI originApp;
        public WorksheetPage euBK;
        public Worksheet euSheet;
        public MatrixPage rlMk;
        public MatrixSheet rlSheet;

        public OutputFile()
        {
            try
            {
                //Create the Origin COM object:
                originApp = new ApplicationCOMSI();
                originApp.Visible = MAINWND_VISIBLE.MAINWND_SHOW;
                originApp.CanClose = true;
                if (originApp == null)
                {
                    Console.WriteLine("Origin could not be started. Check that your installation and project references are correct.");
                    return;
                }
                //Initialize new project:
                originApp.NewProject();
            }
            catch
            {
                Console.WriteLine("ERROR");
            }
        }

        public void EU2Origin(InputFile myInputFile)
        {
            var freq = myInputFile.frequency.Map(x => x.Real).ToArray();
            var e1 = myInputFile.epsilon.Map(x => x.Real).ToArray();
            var e2 = myInputFile.epsilon.Map(x => -x.Imaginary).ToArray();
            var u1 = myInputFile.mu.Map(x => x.Real).ToArray();
            var u2 = myInputFile.mu.Map(x => -x.Imaginary).ToArray();

            try
            {
                //Add a workbook
                euBK = originApp.WorksheetPages.Add(Type.Missing, Type.Missing);
                euBK.Name = myInputFile.FileName + "_EU";
                //The sheet:
                euSheet = (Worksheet)euBK.Layers[0];
                euSheet.Name = "Epsilon and Mu";
                //Add Columns
                for (int i = 0; i < 6; i++)
                {
                    euSheet.Columns.Add(Type.Missing);
                }
                //Set LongName and Units as visible
                euSheet.set_LabelVisible(LABELTYPEVALS.LT_LONG_NAME, true);
                euSheet.set_LabelVisible(LABELTYPEVALS.LT_UNIT, true);
                euSheet.set_LabelVisible(LABELTYPEVALS.LT_COMMENT, true);
                euSheet.set_LabelVisible(LABELTYPEVALS.LT_SPARKLINE, true);
                //Set Long Names, Units, and Comment to the two columns:
                euSheet.Columns[0].LongName = "Frequency";
                euSheet.Columns[0].Units = "GHz";
                euSheet.Columns[1].LongName = "e1"; // permittivity
                euSheet.Columns[2].LongName = "e2";
                euSheet.Columns[3].LongName = "u1"; // permeability
                euSheet.Columns[4].LongName = "u2";
                //Set column types:
                euSheet.Columns[0].Type = COLTYPES.COLTYPE_X;
                for (int i = 1; i < 6; i++)
                {
                    euSheet.Columns[i].Type = COLTYPES.COLTYPE_Y;
                }

                string strForEUName = "[" + euBK.Name + "]" + euSheet.Name;
                originApp.PutWorksheet(strForEUName, freq);
                originApp.PutWorksheet(strForEUName, e1, 0, -1);
                originApp.PutWorksheet(strForEUName, e2, 0, -1);
                originApp.PutWorksheet(strForEUName, u1, 0, -1);
                originApp.PutWorksheet(strForEUName, u2, 0, -1);
            }
            catch (Exception)
            {
                Console.WriteLine("ERROR!!!");
            }
        }

        public void RL2Origin(InputFile myInputFile, Matrix<double> result)
        {
            var nNumRows = result.RowCount;
            var nNumCols = result.ColumnCount;
            var resultRL = result.Transpose().ToArray(); // X - frequency, Y - thickness

            try
            {
                // Create a new matrix book.
                rlMk = originApp.MatrixPages.Add("origin", Type.Missing);
                rlMk.Name = myInputFile.FileName + "_RL";
                // Get the first matrix sheet in the new book.
                rlSheet = (MatrixSheet)rlMk.Layers[0];
                rlSheet.Name = "Reflection Loss";
                rlSheet.Rows = nNumRows;
                rlSheet.Cols = nNumCols;

                // Construct a full sheet name to pass to PutMatrix.
                string srtForRLName = "[" + rlMk.Name + "]" + rlSheet.Name;
                // Put the data into the matrix sheet.
                originApp.PutMatrix(srtForRLName, resultRL);
            }
            catch (Exception)
            {
                Console.WriteLine("ERROR!!!");
            }
        }

        public void ExportToOrigin(InputFile myInputFile, Matrix<double> result)
        {
            this.EU2Origin(myInputFile);
            this.RL2Origin(myInputFile, result);
        }

        public void SaveToOPJ(InputFile myInputFile, string path)
        {
            String PathName = path + "\\" + myInputFile.FileName + ".opj";
            if (originApp.Save(PathName) == false)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("ERROR!!! Failed to save the project into " + PathName);
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("#3/3 Saved into " + PathName + '.');
                Console.ResetColor();
            }
            Console.WriteLine("===========================================================================");
        }
    }


    class Program
    {
        static void Main(string[] args)
        {
            PrintHeader();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("#0/3 Create or select a path to store all the Origin files(.opj), such as \"D:\\eu2rl\\Origin\".");
            Console.ResetColor();
            string storePath = Console.ReadLine();

            do
            {
                InputFile myInputFile = new InputFile();
                myInputFile.ReadCSV();
                var thick = GetThickness();
                var result = myInputFile.CalcRL(thick);
                int fi, di;
                double rlm;
                FindMaxRL(result, out fi, out di, out rlm);
                Console.WriteLine(result);
                Console.WriteLine("Max RL = {0:F2} dB, @ {1:F2} GHz, with {2:F2} mm.", rlm, myInputFile.frequency[fi].Real, thick[di]);

                OutputFile myOutFile = new OutputFile();
                myOutFile.ExportToOrigin(myInputFile, result);
                AddLabel(myInputFile, thick, myOutFile);

                myOutFile.SaveToOPJ(myInputFile, storePath);
            } while (true);
        }

        private static void PrintHeader()
        {
            Console.WriteLine();
            Console.WriteLine("============================== EU2RL TOOLKIT ==============================");
            Console.WriteLine("                     Version 1.1.0 by Shi Kouzhong");
            Console.WriteLine("===========================================================================");
            Console.WriteLine();
        }

        private static Vector<float> GetThickness()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("#2/3 Pelese type the target thickness(mm) to calculate...");
            Console.WriteLine("Case 1： Type several different thickness separated by comma (,), such as 2, 3, 4, ….");
            Console.WriteLine("Case 2： Type a vector by \"begin : step : end\", such as \"2 : 0.01 : 6\".");
            Console.ResetColor();

            List<float> listThick = new List<float>();
            var myInput = Console.ReadLine();
            try
            {
                if (myInput.Contains(','))
                {
                    var myInputValues = myInput.Split(',');
                    foreach (var item in myInputValues)
                    {
                        listThick.Add(float.Parse(item));
                    }
                }
                else if (myInput.Contains(':'))
                {
                    var myInputValues = myInput.Split(':');
                    var begin = double.Parse(myInputValues[0]);
                    var step = double.Parse(myInputValues[1]);
                    var end = double.Parse(myInputValues[2]);
                    var n = Math.Floor((end - begin) / step);
                    for (int i = 0; i <= n; i++)
                    {
                        listThick.Add((float)(begin + i * step));
                    }
                }
                else
                {
                    listThick.Add(float.Parse(myInput));
                }
            }
            catch (Exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR Format!!!");
                Console.ResetColor();
                GetThickness();
            }

            var myThickArray = listThick.ToArray();
            Vector<float> myThick = Vector<float>.Build.Dense(myThickArray);
            return myThick;
        }

        private static void FindMaxRL(Matrix<double> result, out int fi, out int di, out double rlm)
        {
            fi = 0;
            di = 0;
            rlm = 0;
            for (int i = 0; i < result.RowCount; i++)
            {
                for (int j = 0; j < result.ColumnCount; j++)
                {
                    if (result[i, j] < rlm)
                    {
                        rlm = result[i, j];
                        fi = i;
                        di = j;
                    }
                }
            }
        }

        private static void AddLabel(InputFile myInputFile, Vector<float> thick, OutputFile myOutFile)
        {
            var f1 = myInputFile.frequency[0].Real.ToString();
            var f2 = myInputFile.frequency[myInputFile.frequency.Count - 1].Real.ToString();
            var t1 = thick[0].ToString();
            var t2 = thick[thick.Count - 1].ToString();
            var target = "[" + myInputFile.FileName + "RL] Reflection Loss";
            var myS = "mdim x1:=" + f1 + " x2:=" + f2 + " y1:=" + t1 + " y2:=" + t2 + ";"
                      + "range ms = " + target + ";"
                      + "ms.x.longname$ = Frequency;"
                      + "ms.x.units$ = GHz;"
                      + "ms.y.longname$ = Thickness;"
                      + "ms.y.units$ = mm;"
                      + "range mo = 1;"
                      + "mo.label$ = Reflection Loss;"
                      + "mo.unit$ = dB;";
            myOutFile.rlSheet.Execute(myS);
        }
    }
}