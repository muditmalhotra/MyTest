using OfficeOpenXml;
using System.Linq;
using PaySlipEngine.Constant;
using System.Collections.Concurrent;
using PaySlipEngine.Model;
using System.Threading.Tasks;
using OfficeOpenXml.Style;
using PaySlipFactory;
using PaySlipFactory.StateFactories;
using System;

namespace PaySlipGenerator.Helper
{
	public static class PaySlipWorker_COPY
	{
		public static ExcelPackage GeneratePaySlipsExcel(ExcelPackage package)
		{
			int idxLastName = 1, idxFirstName = 0, idxAnualSalary = 2, idxSuperRate = 3, idxPayPeriod = 4;

			//Input
			var workSheet = package.Workbook.Worksheets.First();

			//Output
			var excelExport = new ExcelPackage();
			var workSheetOutput = excelExport.Workbook.Worksheets.Add("Transaction");
			workSheetOutput.TabColor = System.Drawing.Color.Black;
			workSheetOutput.DefaultRowHeight = 12;
			//Header of table  
			workSheetOutput.Row(1).Height = 20;
			workSheetOutput.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			workSheetOutput.Row(1).Style.Font.Bold = true;

			workSheetOutput.Cells[1, 1].Value = "Name";
			workSheetOutput.Cells[1, 2].Value = "GrossIncome";
			workSheetOutput.Cells[1, 3].Value = "IncomeTax";
			workSheetOutput.Cells[1, 4].Value = "NetIncome";
			workSheetOutput.Cells[1, 5].Value = "Super";
			workSheetOutput.Cells[1, 6].Value = "PayPeriod";

			var maxColumnCount = workSheet.Dimension.End.Column;

			Parallel.For(1, maxColumnCount + 1, iCol =>
			{
				var colName = Convert.ToString(((object[,])workSheet.Cells[1, 1, 1, maxColumnCount].Value)[0, iCol - 1]);
				switch (colName)
				{
					case InputExcelColumn.FirstName:
						idxFirstName = iCol;
						break;
					case InputExcelColumn.LastName:
						idxLastName = iCol;
						break;
					case InputExcelColumn.AnnualSalary:
						idxAnualSalary = iCol;
						break;
					case InputExcelColumn.SuperRate:
						idxSuperRate = iCol;
						break;
					case InputExcelColumn.PayPeriod:
						idxPayPeriod = iCol;
						break;
				}
			});


			var blockingCollection = new BlockingCollection<EngineInput>();
			Task.Factory.StartNew(() =>
			{
				EngineInput inputObj;
				Parallel.For(2, maxColumnCount + 1, rowNumber =>
				{
					decimal dVal;
					var row = workSheet.Cells[rowNumber, 1, rowNumber, maxColumnCount];
					inputObj = new EngineInput();

					Parallel.For(1, maxColumnCount + 1, iCol =>
					{
						var value = Convert.ToString(((object[,])row.Value)[0, iCol - 1]);
						if (iCol == idxFirstName)
						{
							inputObj.FirstName = value;
						}
						else if (iCol == idxLastName)
						{
							inputObj.LastName = value;
						}
						else if (iCol == idxAnualSalary)
						{
							inputObj.AnnualSalary = decimal.TryParse(value, out dVal)
								? dVal
								: throw new InvalidCastException(
									$"Annual Salary is not in correct format. Value: {value}");
						}
						else if (iCol == idxSuperRate)
						{
							inputObj.SuperRate = decimal.TryParse(value, out dVal)
								? dVal
								: throw new InvalidCastException($"Super Rate is not in correct format. Value: {value}");
						}
						else if (iCol == idxPayPeriod)
						{
							inputObj.PayPeriod = value;
						}
					});
					blockingCollection.Add(inputObj);
				});

				blockingCollection.CompleteAdding();
			});

			var recordIndex = 2;
			var consumer = Task.Factory.StartNew(() =>
			{
				var state = "NSW";
				PaySlipEngineFactory factory = null;
				switch (state)
				{
					case "NSW":
						factory = new NSWFactory();
						break;
					case "Victoria":
						factory = new VictoriaFactory();
						break;
					default:
						throw new InvalidOperationException("Unknown State.");
				}

				var payEngine = factory.GetPaySlipEngine();

				foreach (var input in blockingCollection.GetConsumingEnumerable())
				{
					var paySlipOutput = payEngine.GeneratePaySlip(input);

					workSheetOutput.Cells[recordIndex, 1].Value = paySlipOutput.Name;
					workSheetOutput.Cells[recordIndex, 2].Value = paySlipOutput.GrossIncome;
					workSheetOutput.Cells[recordIndex, 3].Value = paySlipOutput.IncomeTax;
					workSheetOutput.Cells[recordIndex, 4].Value = paySlipOutput.NetIncome;
					workSheetOutput.Cells[recordIndex, 5].Value = paySlipOutput.Super;
					workSheetOutput.Cells[recordIndex, 6].Value = paySlipOutput.PayPeriod;

					recordIndex++;
				}

			});

			consumer.Wait();
			return excelExport;
		}
	}
}
