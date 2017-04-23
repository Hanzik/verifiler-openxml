using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using VerifilerCore;

namespace VerifilerOpenXML {

	/// <summary>
	/// This validation step is using the Open XML SDK for Office.
	/// 
	/// The error code produced by this validation is Error.Corrupted.
	/// </summary>
	public class XLSXValidator : FormatSpecificValidator {

		public override int ErrorCode { get; set; } = Error.Corrupted;

		public override void Setup() {
			Name = "Microsoft Excel .xlsx files Verification";
			RelevantExtensions.Add(".xlsx");
			Enable();
		}

		public override void ValidateFile(string file) {
			FileStream stream = null;
			try {
				stream = File.Open(file, FileMode.Open);
				SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true);
			} catch (Exception e) {
				ReportAsError("File is corrupted: " + file + "; Message: " + e.Message);
				GC.Collect();
			} finally {
				if (stream != null) {
					stream.Close();
				}
			}
		}
	}
}