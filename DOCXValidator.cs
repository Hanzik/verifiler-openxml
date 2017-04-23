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
	public class DOCXValidator : FormatSpecificValidator {

		public override int ErrorCode { get; set; } = Error.Corrupted;

		public override void Setup() {
			Name = "Microsoft Word files Verification";
			RelevantExtensions.Add(".docx");
			Enable();
		}

		public override void ValidateFile(string file) {
			FileStream stream = null;
			try {
				stream = File.Open(file, FileMode.Open);
				WordprocessingDocument presentation = WordprocessingDocument.Open(stream, true);
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