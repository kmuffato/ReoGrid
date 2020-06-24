/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * http://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Kleber Muffato
 *
 * Copyright (c) 2012-2016 Jing <lujing at unvell.com>
 * Copyright (c) 2012-2016 unvell.com, all rights reserved.
 * 
 ****************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using unvell.Common;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Interaction;
using unvell.ReoGrid.IO;
using unvell.ReoGrid.IO.OpenXML;
using unvell.ReoGrid.IO.OpenXML.Schema;
using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid
{
	partial class Worksheet
	{
		public void LoadGeneric<T>(List<T> obj, string sheetName = "Sheet1")
		{
			this.controlAdapter.ChangeCursor(CursorStyle.Busy); 

			try
			{
				GenericFormatProvider<T> genericProvider = new GenericFormatProvider<T>();

				var arg = new GenericFormatArgument
				{
					 SheetName = sheetName,
					  //Stylesheet = stylesheet,
				};

				Clear();

				genericProvider.Load(this.workbook, obj, arg);
			}
			finally
			{
				this.controlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
			}
		}

		
	}

	internal class GenericFormatProvider<T>
	{
		private const int DEFAULT_READ_BUFFER_ITEMS = 512;

		public void Load(IWorkbook workbook, List<T> obj, object arg)
		{
			bool autoSpread = true;
			string sheetName = String.Empty;
			int bufferItems = DEFAULT_READ_BUFFER_ITEMS > obj.Count ? obj.Count : DEFAULT_READ_BUFFER_ITEMS;
			RangePosition targetRange = RangePosition.EntireRange;

			GenericFormatArgument genericArg = arg as GenericFormatArgument;

			if (genericArg != null)
			{
				sheetName = genericArg.SheetName;
				targetRange = genericArg.TargetRange;
			}

			Worksheet sheet = null;

			if (workbook.Worksheets.Count == 0)
			{
				sheet = workbook.CreateWorksheet(sheetName);
				workbook.Worksheets.Add(sheet);
			}
			else
			{
				while (workbook.Worksheets.Count > 1)
				{
					workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);
				}

				sheet = workbook.Worksheets[0];
				sheet.Reset();
			}

			this.Read(obj, sheet, targetRange, bufferItems, autoSpread);

			// ApplyStyles(sheet, genericArg.Stylesheet);
		}

		private void Read(List<T> obj, Worksheet sheet, RangePosition targetRange, int bufferItems = DEFAULT_READ_BUFFER_ITEMS, bool autoSpread = true)
		{
			targetRange = sheet.FixRange(targetRange);

			Type type = typeof(T);

			T[] items = new T[bufferItems];
			List<object>[] bufferLineList = new List<object>[bufferItems];

			for (int i = 0; i < bufferLineList.Length; i++)
			{
				bufferLineList[i] = new List<object>(256);
			}

#if DEBUG
			var sw = System.Diagnostics.Stopwatch.StartNew();
#endif

			int row = targetRange.Row;
			int totalReadItems = 0;

			sheet.SuspendDataChangedEvents();
			int maxCols = 0;

			try
			{
				bool finished = false;

				while (!finished)
				{
					int readItems = 0;

					for (; readItems < items.Length; readItems++)
					{
						if (totalReadItems >= obj.Count)
						{
							finished = true;
							break;
						}

						items[readItems] = obj[totalReadItems];

						totalReadItems++;
						if (!autoSpread && totalReadItems > targetRange.Rows)
						{
							finished = true;
							break;
						}
					}

					if (autoSpread && row + readItems > sheet.RowCount)
					{
						int appendRows = bufferItems - (sheet.RowCount % bufferItems);
						
						if (appendRows <= 0)
						{
							appendRows = bufferItems;
						}

						sheet.AppendRows(appendRows);
					}

					for (int i = 0; i < readItems; i++)
					{
						var item = items[i];

						var toBuffer = bufferLineList[i];
						toBuffer.Clear();

						foreach (var m in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly).Where(w => w.CanWrite))
						{
							toBuffer.Add(m.GetValue(item, null));

							if (toBuffer.Count >= targetRange.Cols) break;
						}

						if (maxCols < toBuffer.Count) maxCols = toBuffer.Count;

						if (autoSpread && maxCols <= sheet.ColumnCount)
						{
							sheet.SetCols(maxCols + 1);
						}
					}

					sheet.SetRangeData(row, targetRange.Col, readItems, maxCols, bufferLineList);
					row += readItems;
				}

			}
			finally
			{
				sheet.ResumeDataChangedEvents();
			}

			sheet.RaiseRangeDataChangedEvent(new RangePosition(targetRange.Row, targetRange.Col, maxCols, totalReadItems));

#if DEBUG
			sw.Stop();
			System.Diagnostics.Debug.WriteLine("load generic list: " + sw.ElapsedMilliseconds + " ms, rows: " + row);
#endif
		}

	}

	public class GenericFormatArgument
	{
		public string SheetName { get; set; }

		public RangePosition TargetRange { get; set; }

		public Stylesheet Stylesheet { get; set; }
		public GenericFormatArgument()
		{
			this.SheetName = "Sheet1";
			this.TargetRange = RangePosition.EntireRange;
		}
	}


}
