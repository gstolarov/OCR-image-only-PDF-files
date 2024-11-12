using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using System.Net;
namespace ocr {
	public class PdfOcr {
		public static string GhostScript = @"C:\Program Files\gs\gs9.52\bin\gswin64c.exe";
		public static string Tesseract = @"C:\Program Files\Tesseract-OCR\tesseract.exe";
		DirectoryInfo wdiTemp = new DirectoryInfo(Path.GetTempPath() + @"\ocr\");			// temp folder
		string tmpPrfx = DateTime.Now.Ticks.ToString() + "_";								// all temp files will have this prefix									
		public PdfOcr() {
			if (!wdiTemp.Exists)															// make sure temp foilder exists
				Directory.CreateDirectory(wdiTemp.FullName);
			wdiTemp.GetFiles(tmpPrfx + "*.*").ToList().ForEach(x => x.Delete());			// and there is no files like we are going to create
		}
		class Chunk {																		// holds information about text chunk (word/line)
			public float left, top, w, h;													// coordinates converted to points (*72/300)
			public string txt;
			public Chunk(XmlNode n) {
				txt = WebUtility.HtmlDecode(n.InnerText.Trim().Replace("\r", "\n")
							.Replace("Ã¢â‚¬â€", "-").Replace("â€", "-"));
				float[] p = (n.Attributes["title"].Value.Split(new char[] {';'})			// attribute title hold bbox left top right bottom
							.Where(x => x.StartsWith("bbox")).FirstOrDefault() ?? "")		// loat it in props
						.Split(new char[] { ' ' }).Skip(1).Select(x => float.Parse(x) * 72 / 300).ToArray();
				left = p[0]; top = p[1]; h = p[3] - p[1]; w = p[2] - p[0];
			}
		}
		string Hocr2Pdf(iTextSharp.text.pdf.PdfStamper stamp, int pg, string hocr) {
			XmlDocument doc = new XmlDocument();
			string xml = File.ReadAllText(hocr);                                            // get all HOCR file
			xml = Regex.Replace(xml, @"<\?.*\?>|<!DOCTYPE[^>]*>|xmlns=""[^""]*""", "");     // remove DOCTYPE and namespace
			doc.LoadXml(xml);
			iTextSharp.text.pdf.BaseFont fnt = iTextSharp.text.pdf.BaseFont.CreateFont
					(iTextSharp.text.pdf.BaseFont.HELVETICA, iTextSharp.text.pdf.BaseFont.WINANSI, false);
			List<List<Chunk>> lines = new List<List<Chunk>>();
			iTextSharp.text.pdf.PdfGState gState = new iTextSharp.text.pdf.PdfGState();
			gState.FillOpacity = gState.StrokeOpacity = 0;
			foreach (XmlNode ln in doc.SelectNodes("//span[@class='ocrx_word']/..")) {          // ocrx_word parent is a line seqment
				List<Chunk> words = ln.SelectNodes("span[@class='ocrx_word']").OfType<XmlNode>()
					.Select(x => new Chunk(x)).Where(x => x.h <= 50).OrderBy(x => x.left).ToList();
				if (words.Count() < 1) continue;                                                // ignore if height > 50?
				float ht = words.Select(x => x.h).Max(), tp = words.Select(x => x.top).Min();   // all the words in segment same height and base
				words.Sort((x, y) => (int)(x.left - y.left));                                   // order them left to right just in case
				words.ForEach(cw => {
					cw.h = ht; cw.top = tp;
					iTextSharp.text.pdf.PdfContentByte cb = stamp.GetUnderContent(pg);
					cb.SetGState(gState);
					cb.BeginText();
					float size = 8;                                                             // find right font size
					for (; size < 50 && fnt.GetWidthPoint(cw.txt, size) < cw.w; size++)         // where width of the text less then specd
						;
					cb.SetFontAndSize(fnt, cw.h >= 2 ? size - 0.5f : 2);
					cb.SetTextMatrix(cw.left, stamp.Reader.GetPageSizeWithRotation(pg).Height - cw.top - cw.h);
					// + (fnt.GetAscentPoint(cw.txt, size) - fnt.GetDescentPoint(cw.txt, size))/2);
					cb.SetWordSpacing(iTextSharp.text.pdf.PdfWriter.SPACE);
					cb.ShowText(cw.txt + " ");
					cb.EndText();
				});
				lines.Add(words);
			}
			// next portion will not work for multicolumn text...
			lines.Sort((x, y) => (int)(x[0].top - y[0].top));                                   // order lines top to bottom
			for (int i = lines.Count() - 2; i >= 0; i--) {                                      // if we can merge cur with next line
				Chunk c = lines[i][0], n = lines[i + 1][0];                                     // c higher then n ===> c.t < n.t
				if (c.top + c.h > n.top) {
					lines[i].AddRange(lines[i + 1]);                                            // add words from the next line to cur
					lines[i][0].top = c.top;                                                    // reset pos
					lines[i][0].h = Math.Max(c.top + c.h, n.top + n.h) - c.top;
					lines.RemoveAt(i + 1);                                                      // remove next line
				}
			}
			lines.ForEach(r => r.Sort((x, y) => (int)(x.left - y.left)));                       // reorder each line left to right
			string ret = string.Join("\n", lines.Select(r => string.Join(" ", r.Select(c => c.txt).ToList())).ToList());
			return ret.Trim();
		}
		void runProc(string cmd, string arg) {
			List<string> sbErr = new List<string>(), sbOut = new List<string>();
			Process p = new Process() {
				StartInfo = new ProcessStartInfo() {
					FileName = cmd, Arguments = arg,
					WindowStyle = ProcessWindowStyle.Hidden, CreateNoWindow = true,
					UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true
				}
			};
			p.OutputDataReceived += (s, e) => { if (!string.IsNullOrEmpty(e.Data)) sbOut.Add(e.Data); };
			p.ErrorDataReceived += (s, e) => { if (!string.IsNullOrEmpty(e.Data)) sbErr.Add(e.Data); };
			p.Start();
			p.BeginOutputReadLine();
			p.BeginErrorReadLine();
			p.WaitForExit();
			int rc = p.ExitCode;
			if (rc != 0) 
				throw new Exception(string.Format(@"Error {0} to run {1} {2}\n{3}\n{4}", 
					rc, cmd, arg, string.Join("\n", sbErr), string.Join("\n", sbOut)));
		}

		public string OcrFile(string fnIn, string fnOut) {
			if (File.Exists(fnOut)) File.Delete(fnOut);
			wdiTemp.GetFiles(tmpPrfx + "*.*").ToList().ForEach(x => x.Delete());			// make sure temp area is clean for our prefix
			List<string> sb = new List<string>();
			using (iTextSharp.text.pdf.PdfReader rdr = new iTextSharp.text.pdf.PdfReader(fnIn)) {
				int noPages = rdr.NumberOfPages;
				sb.AddRange(Enumerable.Range(1, Math.Min(noPages, 3))
					.Select(i => iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(rdr, i).Trim())
					.Where(i => i != null && i.Length > 0).ToList());
				if (sb.Count() > 0) {
					File.Copy(fnIn, fnOut);
					return string.Join("\n", sb);
				}

				runProc(GhostScript, String.Format(@"-dSAFER -dBATCH -dNOPAUSE -sDEVICE=jpeg -r300 -o ""{0}%04d.jpg"" -f ""{1}""",
											wdiTemp.FullName + tmpPrfx, fnIn));				// run ghostscript to create jpg file in temp
				List<FileInfo> jpegs = wdiTemp.GetFiles(tmpPrfx + "*.jpg").ToList();		// get new jpg files
				using (var cntDown = new CountdownEvent(jpegs.Count())) {					// for each of the jpg run tesseract 
					jpegs.ForEach(f => ThreadPool.QueueUserWorkItem(x => {					// in thread pool to create a HOCR file in temp
						FileInfo fi = (FileInfo)x;
						runProc(Tesseract, String.Format(@"-l eng ""{0}"" ""{1}"" hocr",
								fi.FullName, wdiTemp.FullName + Path.GetFileNameWithoutExtension(fi.Name)));
						cntDown.Signal();													// signal current proc done
						fi.Delete();														// delete jpg file - not need any more
					}, f));
					cntDown.Wait();															// wait for all tesseract procs to finish
				}
				using (FileStream fsOut = new FileStream(fnOut, FileMode.Create, FileAccess.ReadWrite))
				using (iTextSharp.text.pdf.PdfStamper stamp = new iTextSharp.text.pdf.PdfStamper(rdr, fsOut)) {
					stamp.Writer.CompressionLevel = iTextSharp.text.pdf.PdfStream.BEST_COMPRESSION;
					stamp.SetFullCompression();
					rdr.RemoveUnusedObjects();
					stamp.Reader.RemoveUnusedObjects();
					for (int i = 1; i <= jpegs.Count(); i++)
						sb.Add(Hocr2Pdf(stamp, i, string.Format("{2}{0}{1:000#}.hocr", tmpPrfx, i, wdiTemp.FullName)));
					stamp.Close();
					stamp.Reader.Close();
					fsOut.Close();
				}
			}
			wdiTemp.GetFiles(tmpPrfx + "*.hocr").ToList().ForEach(x => x.Delete());			// remove HOCR files 
			return string.Join("\n", sb);
		}
	}
	class Program {
		static void Main(string[] args) {
			if (args.Count() == 2) {
				string txt = new PdfOcr().OcrFile(args[0], args[1]);
				Console.WriteLine(txt);
			}
			else
				Console.WriteLine("Usage ocr fromFile toFile");
			Console.WriteLine("done");
		}
	}
}
