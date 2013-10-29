using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Linq;
using Excel = NetOffice.ExcelApi;
using NetOffice.OfficeApi;
using NetOffice.ExcelApi.Enums;
using LinqTo2dArray;

namespace ExcelExtensions
{

    internal static class Utilities
    {
        internal static void CreateRangeNameObject(this Excel.Workbook wkb, string rangeName, object refersTo)
        {
            var sCleanedRangeName = rangeName.Replace(" ", "_");
            var rangeNames = wkb.Names.Select(n => n.Name);

            if (!rangeNames.Contains(sCleanedRangeName))
            {
                if (refersTo is string)
                {
                    wkb.Names.Add(sCleanedRangeName, Type.Missing, Type.Missing, Type.Missing
                        , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, refersTo);
                }
                else
                {
                    wkb.Names.Add(sCleanedRangeName, refersTo);
                }
            }

        }
    }

    public static class RangeExtenstions
    {

        /// <summary>
        /// Copies the elements of the 2D object array row wise to a new array of the specified element type and length.
        /// </summary>
        /// <typeparam name="TSource">New array type.</typeparam>
        /// <param name="rng">Target range.</param>
        /// <param name="tdefault">Value of default value.</param>
        /// <returns>One-dimensional array of type TSource</returns>
        /// <remarks> Jon Nyman 20130207 </remarks>
        public static T[] ToArray<T>(this Excel.Range rng, T tDefault)
        {
            return rng.ToArray<T>(o => ConvertExcel(o, tDefault));
        }

        private static T ConvertExcel<T>(object o,T tDefault)
        {
            if(o != null && o.GetType() == typeof(T)){
                return (T) o;
            }
            return tDefault;
        }
        
        /// <summary>
        /// Copies the elements of the 2D object array row wise to a new array of the specified element type and length.
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="sDefault">Default string value.</param>
        /// <returns>One-dimensional array of type String</returns>
        /// <remarks> Jon Nyman 20130301 </remarks>
        public static string[] ToArray(this Excel.Range rng, string sDefault)
        {
            return rng.ToArray<string>(o => o.ConvertExcel(sDefault));
        }

        private static string ConvertExcel(this object o, string sDefault)
        {
            if (o != null && o.GetType() != typeof(int))
            {
                return o.ToString();
            }
            return sDefault;
        }

        /// <summary>
        /// Copies the elements of the 2D object array row wise to a new array of the specified element type and length.
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="dteDefault">Default date value.</param>
        /// <returns>One-dimensional array of type Date</returns>
        /// <remarks> Jon Nyman 20130301 </remarks>
        public static DateTime[] ToArray(this Excel.Range rng, DateTime dteDefault)
        {
            return rng.ToArray<DateTime>(o => o.ConvertExcel(dteDefault));
        }   //End ToArray (Date)

        private static DateTime ConvertExcel(this object o, DateTime dteDefault)
        {
            if (o == null) { return dteDefault; }
            if (o.GetType() == typeof(double))
            {
                return DateTime.FromOADate((double) o);
            } else if (o.GetType() == typeof(string)){
                DateTime dte = dteDefault;
                DateTime.TryParse((string) o, out dte);
                return dte;
            } else {
                return dteDefault;
            }
        }
    
        /// <summary>
        /// Count total rows in all areas
        /// </summary>
        /// <param name="rng">Working Range</param>
        /// <returns>Total number of rows.</returns>
        /// <remarks>Jon Nyman 20120924</remarks>
        public static int RowsCount(this Excel.Range rng)
        {

            if ((rng != null))
            {
                int iRowCount = 0;
                foreach (Excel.Range rArea in rng.Areas)
                {
                    iRowCount += rArea.Rows.Count;
                }
                return iRowCount;
            }
            else
            {
                return 0;
            }

        }

       /// <summary>
        /// Parse congruent range into an object by row-object arrays.
        /// </summary>
        /// <typeparam name="TSource">Source type.</typeparam>
        /// <param name="rng">Target range.</param>
        /// <param name="conversion">Function to convert</param>
        /// <returns></returns>
        public static IEnumerable<TSource> Parse<TSource>(this Excel.Range rng, Func<object[], TSource> conversion)
        {
            if (rng.Areas.Count>1)
                throw new ArgumentException("Congruent ranges only allowed.");

            object[,] array = rng.Get2dArrayValue();

            IEnumerable<TSource> cls = array.Parse<TSource>(conversion);

            return cls;
            
        } //End Parse

        /// <summary>
        /// Copies the elements of the 2D object array row wise to a new array of the specified element type and length.
        /// </summary>
        /// <typeparam name="T">New array type.</typeparam>
        /// <param name="rng">Target range.</param>
        /// <param name="conversion">Casting function of new type.</param>
        /// <param name="rowStart">First row index to start.</param>
        /// <param name="columnStart">First column index to start.</param>
        /// <param name="rowCount">Number of rows.</param>
        /// <param name="columnCount">Number of columns</param>
        /// <returns>One-dimensional array of type TSource</returns>
        /// <remarks> Jon Nyman 20130205
        /// Source http://msmvps.com/blogs/jon_skeet/archive/2011/01/02/reimplementing-linq-to-objects-part-24-toarray.aspx </remarks>
        public static T[] ToArray<T>(this Excel.Range rng, Func<object, T> conversion
                                                , int rowStart, int columnStart, int rowCount, int columnCount) 
        {
            rowStart += 1; columnStart += 1; rowCount += 1; columnCount += 1;
            object[,] array = rng.Get2dArrayValue();

            //Make sure values are within range of array.
            if (rowStart < 0 || columnStart < 0 || rowStart > array.GetUpperBound(0) || columnStart > array.GetUpperBound(1) ||
                rowCount < 1 || rowCount + rowStart - 1 > array.GetUpperBound(0) || columnCount < 1 || columnCount + columnStart - 1 > array.GetUpperBound(1))
                throw new System.IndexOutOfRangeException("Start or end values out of range (Parse)");

            return array.ToArray<T>(conversion, rowStart, columnStart, rowCount, columnCount);
            
        } //End ToArray

        /// <summary>
        /// Copies the elements of the 2D object array row wise to a new array of the specified element type and length.
        /// </summary>
        /// <typeparam name="T">New array type.</typeparam>
        /// <param name="rng">Target range.</param>
        /// <param name="conversion">Casting function of new type.</param>
        /// <returns>One-dimensional array of type TSource</returns>
        /// <remarks> Jon Nyman 20130205
        /// Source http://msmvps.com/blogs/jon_skeet/archive/2011/01/02/reimplementing-linq-to-objects-part-24-toarray.aspx </remarks>
        public static T[] ToArray<T>(this Excel.Range rng, Func<object, T> conversion)
        {
            
            object[,] array = rng.Get2dArrayValue();
            return array.ToArray<T>(conversion);

        } //End ToArray

        /// <summary>
        /// Import Data From Excel as 2D zero-based Object Array
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="AsValue">True -> .Value else .Value2 (Default)</param>
        /// <returns>2D object zero-based array</returns>
        /// <remarks>Jon Nyman 121023</remarks>
        public static object[,] To2dArray(this Excel.Range rng, bool AsValue)
        {

	        if (rng.Areas.Count > 1)
		        return rng.ToArrayFromAreas(AsValue);
            
	        object[,] oResult = rng.Get2dArrayValue(AsValue);

		    int iRowUpper = oResult.GetUpperBound(0);
		    int iColumnUpper = oResult.GetUpperBound(1);
            int iRowLower = oResult.GetLowerBound(0);
            int iColumnLower = oResult.GetLowerBound(1);
            object[,] oaResult = new object[iRowUpper - iRowLower + 1, iColumnUpper - iColumnLower + 1];
		    for (int i = iRowLower; i <= iRowUpper; i++) {
			    for (int j = iColumnLower; j <= iColumnUpper; j++) {
				    oaResult[i - iRowLower, j - iColumnLower] = oResult[i, j];
			    }
		    }
		    return oaResult;
	        
        } //End To2dArray

        /// <summary>
        /// Import Data From Excel as 2D zero-based Object Array with Value2
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <returns>2D object zero-based array</returns>
        /// <remarks>Jon Nyman 121023</remarks>
        public static object[,] To2dArray(this Excel.Range rng)
        {

            return rng.To2dArray(false);
            
        } //End To2dArray

        /// <summary>
        /// Loop through areas of range and return single 2d zero-based object array.
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="AsValue">True -> .Value else .Value2 (Default)</param>
        /// <returns>2D object zero-based array</returns>
        private static object[,] ToArrayFromAreas(this Excel.Range rng, bool AsValue)
        {
            Excel.Range rArea = null;
            int iColumnMax = 1;
            object[][,] Objects2D = new object[rng.Areas.Count][,];
            int iAreaCount = 0;

            foreach (Excel.Range rArea_loopVariable in rng.Areas)
            {
                rArea = rArea_loopVariable;
                Objects2D[iAreaCount] = rArea.Get2dArrayValue(AsValue);
                iColumnMax = Math.Max(iColumnMax, Objects2D[iAreaCount].GetUpperBound(1));
                iAreaCount += 1;
            }
            object[,] oaAreaResult = new object[rng.RowsCount(), iColumnMax];
            int iRow = -1;
            for (int i2DArraysIndex = 0; i2DArraysIndex <= iAreaCount - 1; i2DArraysIndex++)
            {
                for (int iRowArea = 1; iRowArea <= Objects2D[i2DArraysIndex].GetUpperBound(0); iRowArea++)
                {
                    iRow += 1;
                    for (int iColumnArea = 1; iColumnArea <= Objects2D[i2DArraysIndex].GetUpperBound(1); iColumnArea++)
                    {
                        oaAreaResult[iRow, iColumnArea - 1] = Objects2D[i2DArraysIndex][iRowArea, iColumnArea];
                    }
                }
            }
            
            return oaAreaResult;
        } //End ToArrayFromAreas

        /// <summary>
        /// Return 2d 0-based or 1-based object array from range
        /// </summary>
        /// <param name="rng">Target range</param>
        /// <param name="asValue">True -> .Value else .Value2 (Default)</param>
        /// <returns>2D object zero-based or one-based array</returns>
        private static object[,] Get2dArrayValue(this Excel.Range rng, bool asValue)
        {

            object resultValue = asValue ? rng.Value : rng.Value2;
            if (resultValue != null && resultValue.GetType().IsArray) {
                return (object[,]) resultValue;
            }else{
                return new object[,] { { resultValue } }; 
            }
            
        } //End GetValueOfRange

        /// <summary>
        /// Return 2d 0-based or 1-based object array from range
        /// </summary>
        /// <param name="rng">Target range</param>
        /// <returns>2D object zero-based or one-based array</returns>
        private static object[,] Get2dArrayValue(this Excel.Range rng)
        {
            return rng.Get2dArrayValue(false);
        }

        /// <summary>
        /// Resize and offset range object.
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="rowSize">Number of rows in new range.</param>
        /// <param name="columnSize">Number of columns in new range.</param>
        /// <param name="rowOffset">Number of rows to offset in new range.</param>
        /// <param name="columnOffset">Number of columns to offset in new range.</param>
        /// <returns>New resized and offset range.</returns>
        /// <remarks>Jon Nyman 20130207</remarks>
        public static Excel.Range ResizeAndOffset(this Excel.Range rng, object rowSize, object columnSize, int rowOffset, int columnOffset)
        {

	        if (!(columnSize is int) || Convert.ToInt32(columnSize) < 1)
		        columnSize = rng.Columns.Count;
	        if (!(rowSize is int) || Convert.ToInt32(rowSize) < 1)
		        rowSize = rng.Rows.Count;
	        return rng.get_Resize(rowSize, columnSize).get_Offset(rowOffset, columnOffset);

        }

        /// <summary>
        /// Export type enumeration.
        /// </summary>
        private enum XlToExcel
        {
            xlFormulas = 1,
            xlValue2 =2,
            xlValue = 3,
            xlFormulaArray
        }

        /// <summary>
        /// Export data to Excel given the type of export.
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="data">2D data to export.</param>
        /// <param name="xlToExcel">Type of export.</param>
        /// <returns>Range where exported to.</returns>
        private static Excel.Range _ToExcel(this Excel.Range rng, object[,] data, XlToExcel xlToExcel)
        {

            if (rng == null)
                return null;

            //Make sure range and 2D object match in size
            Excel.Range rNew = rng.Resize(data.GetUpperBound(0) + 1, data.GetUpperBound(1) + 1);

            //Send to Excel
            if (xlToExcel == XlToExcel.xlValue2)
                rNew.Value2 = data;
            else if (xlToExcel == XlToExcel.xlFormulas)
                rNew.Formula = data;
            else if (xlToExcel == XlToExcel.xlValue)
                rNew.Value = data;
            else if (xlToExcel == XlToExcel.xlFormulaArray)
                rNew.FormulaArray = data;

            return rNew;
        }

        /// <summary>
        /// Export data to Excel given the type of export.
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="data">Data to export.</param>
        /// <param name="xlToExcel">Type of export.</param>
        /// <returns>Range where exported to.</returns>
        private static Excel.Range _ToExcel(this Excel.Range rng, object data, XlToExcel xlToExcel)
        {
            if (rng == null)
                return null;

            if (data.GetType().IsArray)
            {
                if (((Array)data).Rank == 1)
                {
                    object[,] data2D = ((object[])data).To2dArray<object>(o => o);
                    return rng._ToExcel(data2D, xlToExcel);
                }
                else
                {
                    return rng._ToExcel((object[,])data, xlToExcel);
                }
            } //End If IsArray

            //Send to Excel
            if (xlToExcel == XlToExcel.xlValue2)
                rng.Value2 = data;
            else if (xlToExcel == XlToExcel.xlFormulas)
                rng.Formula = data;
            else if (xlToExcel == XlToExcel.xlValue)
                rng.Value = data;
            else if (xlToExcel == XlToExcel.xlFormulaArray)
                rng.FormulaArray = data;

            return rng;

        }

        /// <summary>
        /// Export Data to Excel as Formula
        /// </summary>
        /// <param name="rng">Target Range</param>
        /// <param name="data">Data to export.</param>
        /// <returns>Range where data was exported to.</returns>
        /// <remarks>Jon Nyman 2013-10-09</remarks>
        public static Excel.Range ToExcelFormula(this Excel.Range rng, object data)
        {
            return rng._ToExcel(data, XlToExcel.xlFormulas);
        }

        /// <summary>
        /// Export Data to Excel as Formula
        /// </summary>
        /// <param name="rng">Target Range</param>
        /// <param name="data">Data to export.</param>
        /// <returns>Range where data was exported to.</returns>
        /// <remarks>Jon Nyman 2013-10-09</remarks>
        public static Excel.Range ToExcelFormula(this Excel.Range rng, object[,] data)
        {
            return rng._ToExcel(data, XlToExcel.xlFormulas);
        }

        /// <summary>
        /// Export Data to Excel
        /// </summary>
        /// <param name="rng">Target Range</param>
        /// <param name="data">Data to export.</param>
        /// <returns>Range where data was exported to.</returns>
        /// <remarks>Jon Nyman 121023
        /// 20130205 Convert to C#
        /// </remarks>
        public static Excel.Range ToExcel(this Excel.Range rng, object[,] data)
        {
            return rng._ToExcel(data, XlToExcel.xlValue2);
	    } // End ToExcel
        
        /// <summary>
        /// Export Data to Excel
        /// </summary>
        /// <param name="rng">Target Range</param>
        /// <param name="data">Data to export.</param>
        /// <returns>Range where data was exported to.</returns>
        /// <remarks>Jon Nyman 121023
        /// 20130205 Convert to C#
        /// 2013-10-09 Combined to single private function.</remarks>
        public static Excel.Range ToExcel(this Excel.Range rng, object data)
        {
            return rng._ToExcel(data, XlToExcel.xlValue2);
        } // End ToExcel

        /// <summary>
        /// Export Data to Excel
        /// </summary>
        /// <param name="rng">Target Range</param>
        /// <param name="data">Data to export.</param>
        /// <returns>Range where data was exported to.</returns>
        /// <remarks>Jon Nyman 121023</remarks>
        public static Excel.Range ToExcel(this Excel.Range rng, double[,] data)
        {
            
            if ((rng != null))
            {
                //Make sure range and 2D object match in size
                Excel.Range rNew = rng.Resize(data.GetUpperBound(0) + 1, data.GetUpperBound(1) + 1);
                //Send to Excel
                rNew.Value2 = data;
                return rNew;
            }

            return null;

        } // End ToExcel


        /// <summary>
        /// Get address of range
        /// </summary>
        /// <param name="rng">Target range</param>
        /// <returns>Address of range.</returns>
        public static string _Address(this Excel.Range rng) { return rng.Address; }
        public static string _Address(this Excel.Range rng, bool rowAbsolute) { return rng.Address(rowAbsolute); }
        public static string _Address(this Excel.Range rng, bool rowAbsolute, bool columnAbsolute) { return rng.Address(rowAbsolute, columnAbsolute); }
        public static string _Address(this Excel.Range rng, bool rowAbsolute, bool columnAbsolute, XlReferenceStyle xlReferenceStyle) 
            { return rng.Address(rowAbsolute, columnAbsolute, xlReferenceStyle); }
        public static string _Address(this Excel.Range rng, bool rowAbsolute, bool columnAbsolute, XlReferenceStyle xlReferenceStyle, bool external)
            { return rng.Address(rowAbsolute, columnAbsolute, xlReferenceStyle, external); }
        public static string _Address(this Excel.Range rng, bool rowAbsolute, bool columnAbsolute, XlReferenceStyle xlReferenceStyle, bool external, Excel.Range relativeTo)
            { return rng.Address(rowAbsolute, columnAbsolute, xlReferenceStyle, external, relativeTo); }
        //End Address

        /// <summary>
        /// Create new table if not already created.
        /// </summary>
        /// <param name="rng">Target range.</param>
        /// <param name="tableName">Name of table.</param>
        /// <returns>ListObject (table object)</returns>
        public static Excel.ListObject CreateTable(this Excel.Range rng, string tableName)
        {
            var wks = (Excel.Worksheet) rng.Parent;
            var sCleanedTableName = tableName.Replace(" ", "_");

            var listObjectNames = wks.ListObjects.Select(l => l.Name);

            if (!listObjectNames.Contains(sCleanedTableName))
            {
                var lo = wks.ListObjects.Add(XlListObjectSourceType.xlSrcRange, rng, Type.Missing, XlYesNoGuess.xlYes);
                lo.Name = sCleanedTableName;
                return lo;
            }
            
            return wks.ListObjects[sCleanedTableName];
            
        }

        /// <summary>
        /// Create a workbook range name from range object.
        /// </summary>
        /// <param name="rng">Target range</param>
        /// <param name="rangeName">Name of target range.</param>
        public static void CreateRangeName(this Excel.Range rng, string rangeName)
        {
            var wkb = (Excel.Workbook)((Excel.Worksheet)rng.Parent).Parent;
            wkb.CreateRangeNameObject(rangeName, rng);
        }

        /// <summary>
        /// Create a workbook dynamic range name from range object.
        /// </summary>
        /// <param name="rng">Target range</param>
        /// <param name="rangeName">Name of dynamic target range</param>
        /// <param name="width">Width of named range (1 being the smallest)</param>
        public static void CreateDynamicRangeName(this Excel.Range rng, string rangeName, int excludeCount, int width)
        {
            if (width < 1)
                width = 1;

            var wks = (Excel.Worksheet)rng.Parent;
            var wkb = (Excel.Workbook)wks.Parent;
            var singleRange = rng.Resize(1, 1);
            string firstAddress = singleRange.Address(true, true, XlReferenceStyle.xlR1C1);
            //RefersToR1C1:= "=Categories!R2C1:INDEX(Categories!R1:R1048575,COUNTA(Categories!C1)+1-1,1)"
            string refersTo = string.Format("='{0}'!{1}:INDEX('{0}'!{2},COUNTA('{0}'!C{3})+1-{4},{5})"
                                            , wks.Name
                                            , firstAddress
                                            , singleRange.EntireColumn.Address(true, true, XlReferenceStyle.xlR1C1)
                                            , 1
                                            , excludeCount
                                            , width);
            wkb.CreateRangeNameObject(rangeName, refersTo);
        }

    } //End class RangeExtenstions

    public static class WorksheetExtensions
    {

        /// <summary>
        /// Get range by worksheet, row/column values
        /// </summary>
        /// <param name="wks">Target worksheet</param>
        /// <param name="row">First row (1-base)
        /// 0-Finds end using .Find method
        /// &lt;0 Column to find last row in</param>
        /// <param name="column">First column (1-base)
        /// 0-Finds end using .Find method
        /// &lt;0 row to find last column in</param>
        /// <param name="rowEnd">Last row (1-base)
        /// 0-Finds end using .Find method
        /// &lt;0 Column to find last row in
        /// if iRow &lt;0 then adds to result of iRow</param>
        /// <param name="columnEnd">First column (1-base)
        /// 0-Finds end using .Find method
        /// &lt;0 row to find last column in
        /// if iCol &lt;0 then adds to result of iCol</param>
        /// <returns>Range or Nothing</returns>
        /// <remarks>Jon Nyman www.SpreadsheetBudget.com 7/12/2012</remarks>
        public static Excel.Range Range(this Excel.Worksheet wks, int row, int column, int rowEnd, int columnEnd)
        {
            try
            {
                int iTemp = 0;

                bool bAddRow = false;
                bool bAddCol = false;

                //Determine if first row was set
                if (row < 1)
                {
                    bAddRow = true;
                    row = wks.LastRow(Math.Abs(row)) - 1;
                }
                else
                {
                    row -= 1;
                }

                //Determine if first column was set
                if (column < 1)
                {
                    bAddCol = true;
                    column = wks.LastColumn(Math.Abs(column)) - 1;
                }
                else
                {
                    column -= 1;
                }

                //Determine if the last row was set, if it isn't then get last row number in wks
                if (rowEnd < 1 & !bAddRow)
                {
                    rowEnd = wks.LastRow(Math.Abs(rowEnd)) - 1;
                }
                else if (bAddRow)
                {
                    rowEnd = row + rowEnd;
                }
                else
                {
                    rowEnd -= 1;
                }

                //Determine if the last column was set, if it isn't then get last column number in wks
                if (columnEnd < 1 & !bAddCol)
                {
                    //If lColEnd is negative then get the last column in row number lColEnd
                    columnEnd = wks.LastColumn(Math.Abs(columnEnd)) - 1;
                }
                else if (bAddCol)
                {
                    columnEnd = column + columnEnd;
                }
                else
                {
                    columnEnd -= 1;
                }

                if (row > rowEnd)
                {
                    iTemp = row;
                    row = rowEnd;
                    rowEnd = iTemp;
                }

                if (column > columnEnd)
                {
                    iTemp = column;
                    column = columnEnd;
                    columnEnd = iTemp;
                }

                //Return range, if there is an error then return nothing.		rSource	Nothing	ExcelDna.Integration.ExcelReference

                Excel.Range rResult = wks.Range(wks.Cells[row + 1, column + 1], wks.Cells[rowEnd + 1, columnEnd + 1]);
                return rResult;

            }
            catch
            {
                return null;
            }

        }

        /// <summary>
        /// Find last used row in worksheet.
        /// </summary>
        /// <param name="wks">Target worksheet.</param>
        /// <returns>Returns Integer</returns>
        /// <remarks></remarks>
        public static int LastRow(this Excel.Worksheet wks)
        {

            return wks.Cells.Find
                    ("*", wks.Range("A1"), XlFindLookIn.xlFormulas, XlLookAt.xlWhole, XlRowCol.xlRows, XlSearchDirection.xlPrevious, false, false)
                    .Row;

        }

        /// <summary>
        /// Find last row in specified column.
        /// </summary>
        /// <param name="wks">Target worksheet.</param>
        /// <param name="column">Target column.</param>
        /// <returns>Integer</returns>
        /// <remarks>Jon Nyman 121107
        /// Source: http://www.mrexcel.com/forum/showthread.php?t=74317 4/22/2009</remarks>
        public static int LastRow(this Excel.Worksheet wks, int column)
        {

            if (column == 0)
                return wks.LastRow();

            dynamic xlApp = wks.Application;
            dynamic iTotalRows = xlApp.Rows.Count;
            if ((wks.Cells[iTotalRows, column].Value2 != null))
            {
                return iTotalRows;
            }
            else
            {
                Excel.Range rng = wks.Cells[iTotalRows, column];
                return rng.End(XlDirection.xlUp).Row;
            }

        }

        /// <summary>
        /// Find last used column in worksheet.
        /// </summary>
        /// <param name="wks">Target worksheet.</param>
        /// <returns>Integer</returns>
        /// <remarks>Jon Nyman 121107</remarks>
        public static int LastColumn(this Excel.Worksheet wks)
        {
            //Find the very last column used (use xlValues if blank formulas aren/t used.
            return wks.UsedRange.Cells.Find("*", wks.Range("A1"), XlFindLookIn.xlFormulas, XlLookAt.xlWhole
                                                , XlRowCol.xlColumns, XlSearchDirection.xlPrevious, false, false).Column;
            
        }

        /// <summary>
        /// Find last column in specified column.
        /// </summary>
        /// <param name="wks">Target worksheet.</param>
        /// <param name="row">Target row.</param>
        /// <returns>Integer</returns>
        /// <remarks>Jon Nyman 121107</remarks>
        public static int LastColumn(this Excel.Worksheet wks, int row)
        {

            if (row == 0)
                return wks.LastColumn();

            //Find last used column in specified row.
            //Determine very last column in workbook (XL 2003 or >=XL 2007?).
            dynamic xlApp = wks.Application;
            dynamic iTotalColumns = xlApp.Columns.Count;
            if ((wks.Cells[row, iTotalColumns].Value2 != null))
            {
                //Last cell is not empty so it is the last column.
                return iTotalColumns;
            }
            else
            {
                //Find the last column using the range End method.
                return ((Excel.Range)wks.Cells[row, iTotalColumns]).End(XlDirection.xlToLeft).Column;
            }

        }

        /// <summary>
        /// Return list of list objects on worksheet.
        /// </summary>
        /// <param name="wks">Target worksheet.</param>
        /// <returns>Return list of list objects on worksheet.</returns>
        public static IEnumerable<Excel.ListObject> Tables(this Excel.Worksheet wks)
        {
            foreach (Excel.ListObject listObject in wks.ListObjects)
            {
                yield return listObject;
            }
            yield return null;
        }//End Tables

    } //End class WorksheetExtensions

    

    public static class WorkbookExtensions
    {

        /// <summary>
        /// Create workbook range name from string if doesn/t already exist.
        /// </summary>
        /// <param name="wkb">Target workbook.</param>
        /// <param name="rangeName">Name of range.</param>
        /// <param name="refersTo">String reference formula.</param>
        public static void CreateRangeName(this Excel.Workbook wkb, string rangeName, string refersTo)
        {
            wkb.CreateRangeNameObject(rangeName, refersTo);
        }

        /// <summary>
        /// Checks if workbook is open
        /// </summary>
        /// <param name="wkbs">Target workbook collection</param>
        /// <param name="sFullName">Full name of workbook.</param>
        /// <returns>Workbook if it is open otherwise returns null.</returns>
        public static Excel.Workbook IsOpen(this Excel.Workbooks wkbs, string sFullName)
        {
            foreach (Excel.Workbook wkb in wkbs)
            {
                if (sFullName == wkb.FullName)
                {
                    return wkb;
                }
            }
            return null;
        } //End IsOpen

        /// <summary>
        /// Determine empty workbooks in workbook collection
        /// </summary>
        /// <param name="wkbs">Target workbook collection</param>
        /// <returns>List of workbooks which are empty.</returns>
        public static List<Excel.Workbook> IsEmpty(this Excel.Workbooks wkbs)
        {
            List<Excel.Workbook> lstWkbs = new List<Excel.Workbook> ();
            foreach (Excel.Workbook wkb in wkbs){
                bool bWksAreEmpty = true;
                foreach (Excel.Worksheet wks in wkb.Worksheets){
                    if (wks.UsedRange.Count > 1)
                    {
                        bWksAreEmpty = false;
                    }
                }
                if (bWksAreEmpty && !wkb.Saved)
                {
                    lstWkbs.Add(wkb);
                }
            }
            return lstWkbs;
        } //End IsEmpty

        public static IEnumerable<Excel.Worksheet> Worksheets(this Excel.Sheets wkso)
        {
            foreach (Excel.Worksheet wks in wkso)
            {
                yield return (Excel.Worksheet)wks;
            }
        }


        /// <summary>
        /// Create worksheet or return existing worksheet.
        /// </summary>
        /// <param name="wkb">Target workbook.</param>
        /// <param name="sWorksheetName">Name or codename of worksheet.</param>
        /// <returns>Worksheet</returns>
        /// <remarks>Jon Nyman 121107</remarks>
        public static Excel.Worksheet CreateWorksheet(this Excel.Workbook wkb, string sWorksheetName)
        {
            
            var wks = wkb.Worksheets.Worksheets().WorksheetExists(sWorksheetName);
            if (wks == null){
                wks = (Excel.Worksheet) wkb.Worksheets.Add();
                wks.Name=sWorksheetName;
            }
            return wks;
        }

        /// <summary>
        /// Determine if worksheet exists.
        /// </summary>
        /// <param name="wkb">Target workbook.</param>
        /// <param name="sName">Name or codename of worksheet.</param>
        /// <returns>Worksheet</returns>
        /// <remarks>Jon Nyman 121107</remarks>
        public static Excel.Worksheet WorksheetExists(this IEnumerable<Excel.Worksheet> wkss, string sWorksheetName)
        {
            var wks = wkss.FirstOrDefault(sheet => ((Excel.Worksheet)sheet).Name == sWorksheetName);
            if (wks == null) wks = wkss.FirstOrDefault(sheet => ((Excel.Worksheet)sheet).CodeName == sWorksheetName);
            return wks;
        }

        /// <summary>
        /// Refresh all pivot tables in a workbook.
        /// </summary>
        /// <param name="wkb">Target workbook.</param>
        /// <remarks>Jon Nyman www.SpreadSheetBudget.com July 12, 2012 -> 121107</remarks>
        public static void RefreshPivotTables(this Excel.Workbook wkb)
        {
            foreach(Excel.Worksheet wks in wkb.Worksheets)
            {
                foreach(Excel.PivotTable pt in (Excel.PivotTables) wks.PivotTables())
                {
                    try 
                    {
                        pt.RefreshTable();
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        //Nothing
                    }

                }
            }
        }

    } //End class WorkbookExtensions

    public static class ApplicationExtensions
    {

        public static Excel.Workbook CreateWorkbook(this Excel.Application xlApp, string fileName)
        {
            
            var wkb = xlApp.Workbooks.Add();
            wkb.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return wkb;

        }

    }

}