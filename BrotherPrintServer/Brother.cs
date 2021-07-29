using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using bpac;

namespace BrotherPrintServer
{
	public class PrintData
    {
		public string printer { get; set; }
		public string template { get; set; }
        public int count { get; set; }
		public PrintOptionConstants[] options { get; set; }
		public Dictionary<string, string> fields { get; set; }
    }

	public class PreviewData
    {
		public string template { get; set; }
        public Dictionary<string, string> fields { get; set; }
		public int width { get; set; }
    }

	public class PrinterInfo
    {
        public string name { get; set; }
        public bool online { get; set; }
        public int mediaId { get; set; }
        public string mediaName { get; set; }
        public int tapeLength { get; set; }
    }

	public class TemplateInfo
    {
        public string name { get; set; }
        public string description { get; set; }
        public string medianame { get; set; }
        public int width { get; set; }
        public int length { get; set; }
    }

	class DocumentSetup
    {
        public string Path { get; set; }
        public IDocument Document { get; set; }
        public string Printer { get; set; }
    }


	public enum ErrorCode
    {
		NO_ERROR = 0x0,
		SPECIFIED_FILE_NOT_FOUND = 0x1,
		SPECIFIED_FILE_CANNOT_BE_SAVED = 0x2,
		INVALID_PATH = 0x3,
		SPECIFIED_FILE_CANNOT_BE_OPENED = 0x4,
		SPECIFIED_VALUE_OUT_OF_RANGE = 0x5,
		SPECIFIED_INDEX_INCORRECT = 0x6,
		SPECIFIED_OBJECT_DOES_NOT_EXIST = 0x7,
		SHEET_DOES_NOT_EXIST = 0x8,
		FAILED_TO_PRINT = 0x9,
		OBJECT_DOES_NOT_EXIST = 0xA,
		CURRENT_PRINTER_NOT_SUPPORTED = 0xB,
		NOT_SUPPORTED = 0xC,
		FAILED_TO_CREATE_PRINTER_DC = 0xD,
		TEMPLATE_FILE_NOT_OPEN = 0xE,
		FILE_TYPE_INCORRECT = 0xF,
		INVALID_OBJECT_SPECIFIED = 0x10,
		INVALID_TYPE_SPECIFIED = 0x11,
		INVALID_DATA_SPECIFIED = 0x12,
		INVALID_DATA_TYPE_SPECIFIED = 0x13,
		NOT_LINKED_TO_DOCUMENT = 0x14,
		INSUFFICIENT_MEMORY = 0x16,
		NON_EDITABLE_OBJECT = 0x17,
		INSUFICCIENT_BUFFER = 0x18,
		REFRENCED_MEMORY_IS_NULL = 0x19,
		NO_MEDIA = 0x65,
		END_OF_MEDIA = 0x66,
		TAPE_CUTTER_JAM = 0x67,
		WEAK_BATTERIES = 0x68,
		PRINTER_BUSY = 0x69,
		TURNED_OFF = 0x6A,
		HIGH_VOLTAGE_ADAPTER = 0x6B,
		REPLACE_MEDIA = 0x6C,
		EXPANSION_BUFFER_FULL = 0x6D,
		COMMUNICATION_ERROR = 0x6E,
		COMMUNICATION_BUFFER_FULL = 0x6F,
		COVER_OPEN = 0x70,
		CANCEL_KEY = 0x71,
		MEDIA_END_DETECTION_ERROR = 0x72,
		OFFLINE = 0xC5,
		GENERAL_ERROR_VALUE = 0xC6,
		INVALID_ERROR_VALUE = 0xC7,
		PRINT_WITH_AND_TAPE_WITH_DIFFERENT = 0x12D,
		UNSUPPORTED_BITMAP_FORMAT = 0x12E,
		PRINT_IMAGE_DATA_CANNOT_BE_OPENED = 0x12F,
		RECEIVED_OTHER_DATA_WILE_WAITING_FOR_PRINT_DATA = 0x130,
		INVALID_HEX_STRING = 0x131,
		INVALID_TAPE_WITH = 0x132,
		SETUP_COMPLETED_BLOCK_RANGE_SPECIFIED = 0x133,
		UID_PARAMETER_VALUE_INCORRECT = 0x134,
		ARGUMENT_VALUE_INCORRECT = 0x135,
		SPECIFIED_VALUE_TOO_LARGE = 0x136,
		FAILED_TO_INITIALIZE_USB_DRIVER = 0x137,
		USB_CONNECTION_ERROR = 0x138,
		USB_DATA_TRANSMISSION_ERROR = 0x139,
		USB_RECEPTION_TIMEOUT = 0x13A,
		USB_DATA_RECEPTION_ERROR = 0x13B,
		RECEIVED_DATA_CANNOT_BE_INTERPRETED = 0x141,
		TAG_WRITE_RANGE_INCORRECT = 0x142,
		OPEN_NOT_EXECUTED = 0x143,
		INCORRECT_FLAG_VALUE = 0x144,
		INCORRECT_WRITE_DATA = 0x145,
		INCORRECT_RETRY_COUNT = 0x146,
		SZ_TAPE_NOT_SET = 0x147,
		OPEN_CALLED_WITH_INVALID_ARGUMENT = 0x148,
		SPECIFIED_PRINTER_NOT_AVAILABLE = 0x149,
		FAILED_TO_CREATE_BPAC_OBJECT = 0x14A,
		BPAC_SAVE_ERROR = 0x14D,
		BPAC_EXPORT_ERROR = 0x14E,
		BPAC_GETTEXT_ERROR = 0x14F,
		BPAC_GETFONTINFO_ERROR = 0x150,
		BPAC_SETTEXT_ERROR = 0x151,
		BPAC_SETFRONTINFO_ERROR = 0x152,
		BPAC_SET_BARCODE_DATA_ERROR = 0x153,
		BPAC_REPLACE_IMAGE_FILE_ERROR = 0x154,
		BPAC_OBJECT_NOT_CREATED = 0x155,
		FAILED_TO_INITIALIZE_PORT = 0x168,
		PORT_OPEN = 0x169,
		PORT_COMMUNICATION_FAILURE = 0x16A,
		PORT_DATA_DELIVERY_FAILURE = 0x16B,
		COMMUNICATION_TIMEOUT = 0x16C,
		PORT_RECEIVE_DATA_FAILURE = 0x16D,
		FAILED_TO_CLOSE_PORT = 0x16E,
		INTERNAL_ERROR = 0x186,
		SPECIFIED_RESOLUTION_INCORRECT = 0x01040001,
		SPECIFIED_WIDTH_INCORRECT = 0x01060001,
		SPECIFIED_HEIGHT_INCORRECT = 0x1060002,
		FAILED_TO_CREATE_IMAGE_DATA = 0x01060003,
		FAILED_TO_CONVERT_IMAGE_DATA = 0x01060004,
		SPECIFIED_PRINTER_DOES_NOT_EXIST = 0x01070001,
		FAILED_TO_CHANGE_TO_SPECIFIED_PRINTER = 0x01070002,
		SPECIFIED_MEDIA_ID_NOT_SUPPORTED = 0x01090001,
		NUMBERS_OF_COPIES_INCORRECT = 0x010B0001,
		IMAGEOBJECT_RENDERING_FAILURE = 0x010B0002,
		FAILED_TO_ROTATE_IMAGE = 0x010B0003,
		FAILED_TO_CONVERT_SHEET_RESOLUTION = 0x010B0005,
		FAILED_TO_AQUIRE_TEXT = 0x010F0001,
		FAILED_TO_SET_TEXT = 0x01100001,
		FAILED_TO_AQUIRE_INDEX = 0x01120001,
		FAILED_TO_CREATE_OBJECT_TO_BE_INSERTED = 0x02010001,
		POSITION_CANNOT_BE_INSERTED = 0x02010002,
		FAILED_TO_AQUIRE_TAPE_CODE = 0x03040001,
		FAILED_TO_CREATE_TEMP_PRINT_OBJECT = 0x03040002,
		PRINTER_OFFLINE = 0x03040003,
		NO_SUPPORTED_MEDIA_PRESENT = 0x03070001,
		FAILED_TO_ACCESS_ARRAY = 0x03070002,
		FAILED_TO_AQUIRE_ARRAY_LOWER_BOUND = 0x03070003,
		FAILED_TO_AQUIRE_ARRAY_UPPER_BOUND = 0x03070004,
		FAILED_TO_RELEASE_ARRAY = 0x03070005,
		FAILED_TO_AQUIRE_PRINTER_HANDLE = 0x030A0001,
		FAILED_TO_SET_STRIKEOUT = 0x11070001,
		FAILED_TO_SET_UNDERLINE = 0x11090001,
		INCORRECT_FONT_NAME = 0x12020001,
		FONT_NAME_UPDATE_FAILURE = 0x12020002,
		FONT_BOLD_UPDATE_FAILURE = 0x12040001,
		FONT_EFFECT_UPDATE_FAILURE = 0x12060001,
		FONT_ITALICS_UPDATE_FAILURE = 0x12080001,
		FONT_MAXPOINT_UPDATE_FAILURE = 0x120A0001,
		GET_ATTR_ADDITION_SUBSTRACTION_INVALID = 0x13010001,
		SET_ATTR_ADDITION_SUBSTRACTION_INVALID = 0x13020001,
		IMAGE_DATA_ARRAY_CREATION_FAILURE = 0x16010001,
		FAILED_TO_ACCESS_IMAGE_DATA = 0x16010002,
		FAILED_TO_RELEASE_IMAGE_DATA_MEMORY = 0x16010003,
		FAILURE_IN_IMAGE_SWITCHING = 0x16020001
	}
	public class ComDisposer : IDisposable
	{
		private List<System.Object> _comObjs;

		public ComDisposer()
		{
			_comObjs = new List<System.Object>();
		}

		~ComDisposer()
		{
			Dispose(false);
		}

		public T Add<T>(T o)
		{
			if ( o != null && o.GetType().IsCOMObject )
				_comObjs.Add(o);
			return o;
		}

		public void Clear()
		{
			Dispose(true);
		}

		protected virtual void Dispose(bool disposing)
		{
			if ( disposing )
			{
				for ( int i = _comObjs.Count - 1; i >= 0; --i )
					Marshal.FinalReleaseComObject( _comObjs[i] );
				_comObjs.Clear();
			}
		}

		void IDisposable.Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
	}
	public class Brother
	{
        static SemaphoreSlim semaphoreSlim = new SemaphoreSlim(1,1);
		private const string TEMPLATE_DIRECTORY = @".\templates\"; // Template file path

		private static List<PrinterInfo> printerInfos;
		private static List<TemplateInfo> templateInfos;

		public static async Task<IEnumerable<PrinterInfo>> GetPrintersAsync(bool forceUpdate = false)
		{
		    await semaphoreSlim.WaitAsync();

				try
				{
			if (printerInfos == null || forceUpdate)
			{
					using (ComDisposer cd = new ComDisposer())
					{
						printerInfos = new List<PrinterInfo>();
						IPrinter printerObject = cd.Add(new Printer());

						object[] printers = printerObject.GetInstalledPrinters();

						if (printers != null)
						{
							foreach (object printer in printers)
							{
								string printerName = printer.ToString();
								Document doc = cd.Add(new Document());
								doc.SetPrinter(printerName, false);

								printerInfos.Add(new PrinterInfo
								{
									name = printerName,
									online = doc.Printer.IsPrinterOnline(printerName),
									mediaId = doc.Printer.GetMediaId(),
									mediaName = doc.Printer.GetMediaName(),
									tapeLength = doc.Printer.GetPrintedTapeLength()
								});
							}
						}
					}
				}
			}
				finally
				{
					semaphoreSlim.Release();
				}
			return printerInfos;
		}

		public static async Task<IEnumerable<TemplateInfo>> GetTemplatesAsync(bool forceUpdate = false)
        {

			await semaphoreSlim.WaitAsync();

			try
			{
				if (templateInfos == null || forceUpdate)
				{
					templateInfos = new List<TemplateInfo>();
					var files = Directory.GetFiles(TEMPLATE_DIRECTORY);
						foreach (var file in files)
						{

					        using (ComDisposer cd = new ComDisposer())
					        {
						    Document doc = cd.Add(new Document());
							if (Path.GetExtension(file).ToLower() == ".lbx")
							{
								if (doc.Open(file))
								{
									templateInfos.Add(new TemplateInfo
									{
										name = Path.GetFileNameWithoutExtension(file),
										width = doc.Width,
										length = doc.Length,
										medianame = doc.GetMediaName()
									});
									doc.Close();
								}
							}
						}
					}
				}
			}

			finally
			{
				semaphoreSlim.Release();
			}
			return templateInfos;
		}

		public static async Task UpdateAsync()
        {
			await GetPrintersAsync(true);
			await GetTemplatesAsync(true);
        }

    	public static async Task<string> PreviewAsync(PreviewData printData)
		{
			string outString = "";
			IDocument doc = null; 
			bool printerOpened = false;
			//sanity check
			if (printData.template == null || printData.template == "")
				return outString;

			await semaphoreSlim.WaitAsync();

			try {
			string templatePath = String.Format("{0}{1}.lbx", TEMPLATE_DIRECTORY, printData.template);


				//using (ComDisposer cd = new ComDisposer())
				//{

					doc = new Document();
					//doc = cd.Add(new Document());

					if (doc.Open(templatePath))
					{
						printerOpened = true;

						foreach (IObject obj in doc.Objects)
						{
							var name = obj.Name.ToLower();
							string key = printData.fields.Keys.FirstOrDefault(m => m.ToLower() == name);
							if (key != null)
								obj.Text = printData.fields[key];
							else
								obj.Text = "";
						}

					//PrintOptionConstants options = PrintOptionConstants.bpoHalfCut | PrintOptionConstants.bpoHighResolution;
					using (var ms = new MemoryStream(doc.GetImageData(ExportType.bexBmp, printData.width, 0)))
					{
						Bitmap bmp = new Bitmap(ms);
						using (var ms_out = new MemoryStream())
						{
							bmp.Save(ms_out, System.Drawing.Imaging.ImageFormat.Png);
							//outString = "data:image/png;base64," + Convert.ToBase64String(ms_out.ToArray());
							outString = Convert.ToBase64String(ms_out.ToArray());
						}

						doc.Close();
                    }
					}
					else
					{
						//MessageBox.Show("Open() Error: " + doc.ErrorCode);
					}
				//}
			 }
			catch(Exception ex)
            {
				Console.WriteLine(ex);
            }
			finally
            {
				if (printerOpened)
					doc.Close();
				semaphoreSlim.Release();
            }
			return outString;
		}
    	public static async Task PrintAsync(PrintData printData)
		{
			IDocument doc = null; 
				bool printerOpened = false;
			//sanity check
			if (printData.template == null || printData.template == "")
				return;

			if (printData.count < 1)
				printData.count = 1;

			await semaphoreSlim.WaitAsync();

			try {
			string templatePath = String.Format("{0}{1}.lbx", TEMPLATE_DIRECTORY, printData.template);

			ErrorCode errorCode;

				//using (ComDisposer cd = new ComDisposer())
				//{

					doc = new Document();
					//doc = cd.Add(new Document());

					doc.SetPrinter(printData.printer, false);
					if (doc.Open(templatePath))
					{
						printerOpened = true;

						foreach (IObject obj in doc.Objects)
						{
							var name = obj.Name.ToLower();
							string key = printData.fields.Keys.FirstOrDefault(m => m.ToLower() == name);
							if (key != null)
								obj.Text = printData.fields[key];
							else
								obj.Text = "";
						}

					//foreach (var nv in printData.fields)
					//{
					//	var obj = documentSetup.Document.GetObject(nv.Key);
					//	if (obj != null) 
					//	  obj.Text = nv.Value;
					//}

					// doc.SetMediaById(doc.Printer.GetMediaId(), true);
					PrintOptionConstants options = PrintOptionConstants.bpoDefault;
					if (printData.options != null)
                    {
						foreach (var opt in printData.options)
							options |= opt;
                    } else
                    {
						options = PrintOptionConstants.bpoHalfCut | PrintOptionConstants.bpoHighResolution;
					}
					
						doc.StartPrint("", options);
						doc.PrintOut(printData.count, options);
						doc.EndPrint();
						doc.Close();
					}
					else
					{
						//MessageBox.Show("Open() Error: " + doc.ErrorCode);
					}
				//}
			 }
			catch(Exception ex)
            {
				Console.WriteLine(ex);
            }
			finally
            {
				if (printerOpened)
					doc.Close();
				semaphoreSlim.Release();
            }
		}
	}
}
