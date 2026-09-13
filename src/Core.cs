using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: RuntimeCompatibility(WrapNonExceptionThrows = true)]
[assembly: AssemblyTitle("Lab Photo Tools")]
[assembly: ComVisible(false)]
[assembly: CompilationRelaxations(8)]
[assembly: AssemblyVersion("0.1.15.0")]
namespace LabPhotoTools
{
	[ComVisible(true)]
	[Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IDTExtensibility2
	{
		[DispId(1)]
		void OnConnection([In][MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode, [In][MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[DispId(2)]
		void OnDisconnection([In] int removeMode, [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[DispId(3)]
		void OnAddInsUpdate([In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[DispId(4)]
		void OnStartupComplete([In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

		[DispId(5)]
		void OnBeginShutdown([In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
	}
	[Guid("000C0396-0000-0000-C000-000000000046")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	[ComVisible(true)]
	public interface IRibbonExtensibility
	{
		[DispId(1)]
		[return: MarshalAs(UnmanagedType.BStr)]
		string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
	}
	internal class WindowOwner : IWin32Window
	{
		public IntPtr Handle { get; private set; }

		public WindowOwner(IntPtr handle)
		{
			Handle = handle;
		}
	}
	public class EngineItem
	{
		public string input { get; set; }

		public string output { get; set; }

		public double angle { get; set; }
	}
	public static class Engine
	{
		public static string Root
		{
			get
			{
				return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			}
		}

		public static string Home
		{
			get
			{
				if (!Directory.Exists(Path.Combine(Root, "engine")))
				{
					return Directory.GetParent(Root).FullName;
				}
				return Root;
			}
		}

		public static string PythonPath
		{
			get
			{
				Dictionary<string, object> dictionary = Config();
				string text = (dictionary.ContainsKey("pythonPath") ? Convert.ToString(dictionary["pythonPath"]) : Path.Combine(Home, "runtime", "Scripts", "python.exe"));
				if (!Path.IsPathRooted(text) || !File.Exists(text))
				{
					throw new InvalidOperationException("사진 처리 도구가 아직 설치되지 않았습니다. 배포 폴더의 설치 안내에 따라 Python 환경을 준비해 주세요.");
				}
				return text;
			}
		}

		private static Dictionary<string, object> Config()
		{
			string path = Path.Combine(Home, "config.json");
			if (!File.Exists(path))
			{
				return new Dictionary<string, object>();
			}
			return new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(path));
		}

		public static string NewJobDirectory()
		{
			string text = Path.Combine(Path.GetTempPath(), "LabPhotoTools", Guid.NewGuid().ToString("N"));
			Directory.CreateDirectory(text);
			return text;
		}

		public static void CleanJob(string path)
		{
			if (string.IsNullOrEmpty(path))
			{
				return;
			}
			string value = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "LabPhotoTools")) + Path.DirectorySeparatorChar;
			string fullPath = Path.GetFullPath(path);
			if (!fullPath.StartsWith(value, StringComparison.OrdinalIgnoreCase) || Path.GetFileName(fullPath).Length != 32)
			{
				return;
			}
			try
			{
				Directory.Delete(fullPath, true);
			}
			catch (IOException)
			{
			}
			catch (UnauthorizedAccessException)
			{
			}
		}

		public static void ValidateReady()
		{
			string pythonPath = PythonPath;
		}

		public static async Task RunAsync(string operation, List<EngineItem> items, string jobDirectory, CancellationToken token)
		{
			string python = PythonPath;
			string worker = Path.Combine(Home, "engine", "worker.py");
			if (!File.Exists(worker))
			{
				throw new InvalidOperationException("사진 처리 파일을 찾을 수 없습니다. 추가 기능을 다시 설치해 주세요.");
			}
			Dictionary<string, object> cfg = Config();
			string modelDir = (cfg.ContainsKey("modelDir") ? Convert.ToString(cfg["modelDir"]) : Path.Combine(Home, "models"));
			string request = Path.Combine(jobDirectory, "request.json");
			File.WriteAllText(request, new JavaScriptSerializer().Serialize(new
			{
				operation = operation,
				items = items,
				model_dir = modelDir
			}), new UTF8Encoding(false));
			await Task.Factory.StartNew(delegate
			{
				ProcessStartInfo startInfo = new ProcessStartInfo(python, "-X utf8 " + Quote(worker) + " --request " + Quote(request))
				{
					UseShellExecute = false,
					CreateNoWindow = true,
					RedirectStandardOutput = true,
					RedirectStandardError = true,
					StandardOutputEncoding = Encoding.UTF8,
					StandardErrorEncoding = Encoding.UTF8,
					WorkingDirectory = Home
				};
				using (Process process = new Process
				{
					StartInfo = startInfo
				})
				{
					process.Start();
					Task<string> task = process.StandardOutput.ReadToEndAsync();
					Task<string> task2 = process.StandardError.ReadToEndAsync();
					Stopwatch stopwatch = Stopwatch.StartNew();
					while (!process.WaitForExit(100))
					{
						if (token.IsCancellationRequested || stopwatch.Elapsed.TotalMinutes > 15.0)
						{
							try
							{
								process.Kill();
							}
							catch (InvalidOperationException)
							{
							}
							process.WaitForExit();
							token.ThrowIfCancellationRequested();
							throw new InvalidOperationException("사진 처리가 15분을 넘었습니다. 사진 수를 줄여 다시 시도하세요.");
						}
					}
					string result = task.Result;
					string result2 = task2.Result;
					token.ThrowIfCancellationRequested();
					Dictionary<string, object> dictionary = null;
					try
					{
						dictionary = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(result);
					}
					catch (ArgumentException)
					{
					}
					if (process.ExitCode != 0 || dictionary == null || !dictionary.ContainsKey("ok") || !Convert.ToBoolean(dictionary["ok"]))
					{
						string message = "사진 처리를 완료하지 못했습니다. Python 환경과 모델 설치 상태를 확인해 주세요.";
						if (dictionary != null && dictionary.ContainsKey("error"))
						{
							Dictionary<string, object> dictionary2 = dictionary["error"] as Dictionary<string, object>;
							if (dictionary2 != null && dictionary2.ContainsKey("message"))
							{
								message = Convert.ToString(dictionary2["message"]);
							}
							else if (dictionary["error"] is string)
							{
								message = (string)dictionary["error"];
							}
						}
						throw new InvalidOperationException(message);
					}
					foreach (EngineItem item in items)
					{
						if (!File.Exists(item.output) || new FileInfo(item.output).Length == 0)
						{
							throw new InvalidOperationException("사진 처리 결과가 누락되었습니다.");
						}
					}
				}
			}, token, TaskCreationOptions.None, TaskScheduler.Default);
		}

		private static string Quote(string value)
		{
			StringBuilder stringBuilder = new StringBuilder("\"");
			int num = 0;
			foreach (char c in value)
			{
				switch (c)
				{
				case '\\':
					num++;
					continue;
				case '"':
					stringBuilder.Append('\\', num * 2 + 1);
					stringBuilder.Append(c);
					break;
				default:
					stringBuilder.Append('\\', num);
					stringBuilder.Append(c);
					break;
				}
				num = 0;
			}
			stringBuilder.Append('\\', num * 2);
			stringBuilder.Append('"');
			return stringBuilder.ToString();
		}
	}
	public class PhotoBox
	{
		public int Id;

		public double Left;

		public double Top;

		public double Width;

		public double Height;
	}
	public static partial class Layout
	{
		private sealed class Row
		{
			public readonly List<PhotoBox> Photos = new List<PhotoBox>();

			public double Top;

			public double MinHeight;
		}

		private const int MaxPhotos = 1000;

		private const double BoundsEpsilon = 1E-06;

		public static List<PhotoBox> Arrange(IList<PhotoBox> photos, int columns, double gapX, double gapY, bool uniformWidth, double targetWidth, double slideWidth, double slideHeight)
		{
			List<PhotoBox> list = ValidateAndClone(photos, gapX, gapY, slideWidth, slideHeight);
			if (columns < 1 || columns > 1000)
			{
				throw new ArgumentException("열 개수는 1~1000 사이여야 합니다.", "columns");
			}
			if (uniformWidth && (!Finite(targetWidth) || targetWidth <= 0.0))
			{
				throw new ArgumentException("통일할 사진 너비는 0보다 큰 숫자여야 합니다.", "targetWidth");
			}
			double num = MinimumLeft(list);
			double num2 = MinimumTop(list);
			List<Row> list2 = DetectRows(list);
			List<PhotoBox> list3 = new List<PhotoBox>(list.Count);
			foreach (Row item in list2)
			{
				list3.AddRange(item.Photos);
			}
			double num3 = num2;
			for (int i = 0; i < list3.Count; i += columns)
			{
				double num4 = num;
				double num5 = 0.0;
				int num6 = Math.Min(i + columns, list3.Count);
				for (int j = i; j < num6; j++)
				{
					PhotoBox photoBox = list3[j];
					if (uniformWidth)
					{
						photoBox.Height *= targetWidth / photoBox.Width;
						photoBox.Width = targetWidth;
					}
					photoBox.Left = num4;
					photoBox.Top = num3;
					num4 += photoBox.Width + gapX;
					num5 = Math.Max(num5, photoBox.Height);
				}
				num3 += num5 + gapY;
			}
			ValidateBounds(list3, slideWidth, slideHeight);
			return list3;
		}

		public static List<PhotoBox> Space(IList<PhotoBox> photos, double gapX, double gapY, bool horizontal, bool vertical, double slideWidth, double slideHeight)
		{
			List<PhotoBox> list = ValidateAndClone(photos, gapX, gapY, slideWidth, slideHeight);
			List<Row> list2 = DetectRows(list);
			List<PhotoBox> list3 = new List<PhotoBox>(list.Count);
			double num = list2[0].Top;
			foreach (Row item in list2)
			{
				double num2 = MinimumLeft(item.Photos);
				double num3 = (vertical ? (num - item.Top) : 0.0);
				double num4 = double.NegativeInfinity;
				foreach (PhotoBox photo in item.Photos)
				{
					if (horizontal)
					{
						photo.Left = num2;
						num2 += photo.Width + gapX;
					}
					photo.Top += num3;
					num4 = Math.Max(num4, photo.Top + photo.Height);
					list3.Add(photo);
				}
				if (vertical)
				{
					num = num4 + gapY;
				}
			}
			ValidateBounds(list3, slideWidth, slideHeight);
			return list3;
		}

		private static List<PhotoBox> ValidateAndClone(IList<PhotoBox> photos, double gapX, double gapY, double slideWidth, double slideHeight)
		{
			if (photos == null || photos.Count == 0)
			{
				throw new ArgumentException("배치할 사진을 한 장 이상 선택하세요.", "photos");
			}
			if (photos.Count > 1000)
			{
				throw new ArgumentException("한 번에 최대 1000장의 사진을 처리할 수 있습니다.", "photos");
			}
			if (!Finite(gapX) || !Finite(gapY) || gapX < 0.0 || gapY < 0.0)
			{
				throw new ArgumentException("사진 간격은 0 이상의 숫자여야 합니다.");
			}
			if (!Finite(slideWidth) || !Finite(slideHeight) || slideWidth <= 0.0 || slideHeight <= 0.0)
			{
				throw new ArgumentException("슬라이드의 너비와 높이를 확인할 수 없습니다.");
			}
			HashSet<int> hashSet = new HashSet<int>();
			List<PhotoBox> list = new List<PhotoBox>(photos.Count);
			foreach (PhotoBox photo in photos)
			{
				if (photo == null || !Finite(photo.Left) || !Finite(photo.Top) || !Finite(photo.Width) || !Finite(photo.Height) || photo.Width <= 0.0 || photo.Height <= 0.0)
				{
					throw new ArgumentException("사진의 위치 또는 크기가 올바르지 않습니다.", "photos");
				}
				if (!hashSet.Add(photo.Id))
				{
					throw new ArgumentException("선택한 사진 목록에 중복된 항목이 있습니다.", "photos");
				}
				list.Add(new PhotoBox
				{
					Id = photo.Id,
					Left = photo.Left,
					Top = photo.Top,
					Width = photo.Width,
					Height = photo.Height
				});
			}
			return list;
		}

		private static List<Row> DetectRows(List<PhotoBox> photos)
		{
			List<PhotoBox> list = new List<PhotoBox>(photos);
			list.Sort(CompareTop);
			List<Row> list2 = new List<Row>();
			foreach (PhotoBox item in list)
			{
				Row row = ((list2.Count == 0) ? null : list2[list2.Count - 1]);
				double num = ((row == null) ? 0.0 : Math.Min(12.0, Math.Min(row.MinHeight, item.Height) * 0.2));
				if (row == null || item.Top - row.Top > num)
				{
					Row row2 = new Row();
					row2.Top = item.Top;
					row2.MinHeight = item.Height;
					row = row2;
					list2.Add(row);
				}
				row.MinHeight = Math.Min(row.MinHeight, item.Height);
				row.Photos.Add(item);
			}
			foreach (Row item2 in list2)
			{
				item2.Photos.Sort(CompareLeft);
			}
			return list2;
		}

		private static int CompareTop(PhotoBox a, PhotoBox b)
		{
			int num = a.Top.CompareTo(b.Top);
			if (num == 0)
			{
				num = a.Left.CompareTo(b.Left);
			}
			if (num != 0)
			{
				return num;
			}
			return a.Id.CompareTo(b.Id);
		}

		private static int CompareLeft(PhotoBox a, PhotoBox b)
		{
			int num = a.Left.CompareTo(b.Left);
			if (num == 0)
			{
				num = a.Top.CompareTo(b.Top);
			}
			if (num != 0)
			{
				return num;
			}
			return a.Id.CompareTo(b.Id);
		}

		private static double MinimumLeft(IList<PhotoBox> photos)
		{
			double num = double.PositiveInfinity;
			foreach (PhotoBox photo in photos)
			{
				num = Math.Min(num, photo.Left);
			}
			return num;
		}

		private static double MinimumTop(IList<PhotoBox> photos)
		{
			double num = double.PositiveInfinity;
			foreach (PhotoBox photo in photos)
			{
				num = Math.Min(num, photo.Top);
			}
			return num;
		}

		private static bool Finite(double value)
		{
			if (!double.IsNaN(value))
			{
				return !double.IsInfinity(value);
			}
			return false;
		}

		private static void ValidateBounds(List<PhotoBox> result, double slideWidth, double slideHeight)
		{
			foreach (PhotoBox item in result)
			{
				if (!Finite(item.Left) || !Finite(item.Top) || !Finite(item.Width) || !Finite(item.Height) || !Finite(item.Left + item.Width) || !Finite(item.Top + item.Height) || item.Width <= 0.0 || item.Height <= 0.0 || item.Left < -1E-06 || item.Top < -1E-06 || item.Left + item.Width > slideWidth + 1E-06 || item.Top + item.Height > slideHeight + 1E-06)
				{
					throw new ArgumentException("배치 결과가 슬라이드 밖으로 나갑니다. 사진 크기, 열 개수 또는 간격을 줄이거나 시작 위치를 조정하세요.");
				}
			}
		}
	}
	internal sealed class ImagePreviewSession : IDisposable
	{
		private readonly string job;

		private readonly CancellationTokenSource cancellation = new CancellationTokenSource();

		private Task backgroundTask;

		private bool disposed;

		public Bitmap Original { get; private set; }

		public Bitmap BackgroundRemoved { get; private set; }

		public ImagePreviewSession(string directory)
		{
			job = directory;
			Original = LoadThumbnail(Path.Combine(job, "input.png"));
		}

		internal static Bitmap LoadThumbnail(string path)
		{
			using (Image image = Image.FromFile(path))
			{
				double num = Math.Min(1.0, 1600.0 / (double)Math.Max(image.Width, image.Height));
				Bitmap bitmap = new Bitmap(Math.Max(1, (int)Math.Round((double)image.Width * num)), Math.Max(1, (int)Math.Round((double)image.Height * num)), PixelFormat.Format32bppPArgb);
				using (Graphics graphics = Graphics.FromImage(bitmap))
				{
					using (ImageAttributes imageAttributes = new ImageAttributes())
					{
						imageAttributes.SetWrapMode(WrapMode.TileFlipXY);
						graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
						graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
						graphics.DrawImage(image, new Rectangle(Point.Empty, bitmap.Size), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, imageAttributes);
					}
				}
				return bitmap;
			}
		}

		public Task PrepareBackgroundAsync()
		{
			if (disposed)
			{
				throw new ObjectDisposedException("ImagePreviewSession");
			}
			if (backgroundTask == null || backgroundTask.IsFaulted || backgroundTask.IsCanceled)
			{
				backgroundTask = PrepareBackgroundCoreAsync();
			}
			return backgroundTask;
		}

		private async Task PrepareBackgroundCoreAsync()
		{
			string output = Path.Combine(job, "background-" + Guid.NewGuid().ToString("N") + ".png");
			try
			{
				await Engine.RunAsync("remove_background", new List<EngineItem>
				{
					new EngineItem
					{
						input = Path.Combine(job, "input.png"),
						output = output,
						angle = 0.0
					}
				}, job, cancellation.Token);
				if (!disposed)
				{
					BackgroundRemoved = LoadThumbnail(output);
				}
			}
			finally
			{
				if (disposed)
				{
					Engine.CleanJob(job);
					cancellation.Dispose();
				}
			}
		}

		public void Dispose()
		{
			if (!disposed)
			{
				disposed = true;
				cancellation.Cancel();
				Original.Dispose();
				if (BackgroundRemoved != null)
				{
					BackgroundRemoved.Dispose();
				}
				if (backgroundTask == null || backgroundTask.IsCompleted)
				{
					Engine.CleanJob(job);
					cancellation.Dispose();
				}
			}
		}
	}
	internal static class PreviewGeometry
	{
		public static double CropScale(double width, double height, double angle)
		{
			double num = angle * Math.PI / 180.0;
			double num2 = Math.Abs(Math.Cos(num));
			double num3 = Math.Abs(Math.Sin(num));
			return Math.Min(width / (width * num2 + height * num3), height / (width * num3 + height * num2));
		}

		public static RectangleF Fit(Size viewport, Size image, double angle, bool crop)
		{
			double num = angle * Math.PI / 180.0;
			double num2 = Math.Abs(Math.Cos(num));
			double num3 = Math.Abs(Math.Sin(num));
			double num4 = (crop ? ((double)image.Width) : ((double)image.Width * num2 + (double)image.Height * num3));
			double num5 = (crop ? ((double)image.Height) : ((double)image.Width * num3 + (double)image.Height * num2));
			double num6 = Math.Min((double)Math.Max(1, viewport.Width - 32) / num4, (double)Math.Max(1, viewport.Height - 32) / num5);
			return new RectangleF((float)(((double)viewport.Width - num4 * num6) / 2.0), (float)(((double)viewport.Height - num5 * num6) / 2.0), (float)(num4 * num6), (float)(num5 * num6));
		}
	}
	internal sealed class PhotoPreview : Control
	{
		public Image Source { get; set; }

		public double Angle { get; set; }

		public bool Crop { get; set; }

		public bool ShowGrid { get; set; }

		public int GridColumns { get; set; }

		public int GridRows { get; set; }

		public PhotoPreview()
		{
			DoubleBuffered = true;
			base.ResizeRedraw = true;
			BackColor = Color.FromArgb(222, 227, 233);
			GridColumns = (GridRows = 4);
			ShowGrid = (Crop = true);
			base.AccessibleName = "첫 사진 자동 미리보기";
			base.TabStop = false;
		}

		protected override void OnPaint(PaintEventArgs e)
		{
			base.OnPaint(e);
			if (Source == null)
			{
				TextRenderer.DrawText(e.Graphics, "사진 미리보기를 준비하고 있습니다…", Font, base.ClientRectangle, Color.FromArgb(70, 85, 102), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
			}
			else
			{
				DrawPhoto(e.Graphics, base.ClientSize, Source, Angle, Crop, ShowGrid, GridColumns, GridRows);
			}
		}

		internal static void DrawPhoto(Graphics g, Size viewport, Image source, double angle, bool crop, bool grid, int columns, int rows)
		{
			RectangleF rect = PreviewGeometry.Fit(viewport, source.Size, angle, crop);
			GraphicsState gstate = g.Save();
			try
			{
				g.SetClip(rect, CombineMode.Intersect);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(248, 249, 251)))
				{
					using (SolidBrush brush2 = new SolidBrush(Color.FromArgb(231, 235, 240)))
					{
						g.FillRectangle(brush, rect);
						for (int i = 0; (float)i < rect.Height / 12f; i++)
						{
							for (int j = 0; (float)j < rect.Width / 12f; j++)
							{
								if ((j + i) % 2 == 0)
								{
									g.FillRectangle(brush2, rect.X + (float)(j * 12), rect.Y + (float)(i * 12), 12f, 12f);
								}
							}
						}
					}
				}
				GraphicsState gstate2 = g.Save();
				try
				{
					double num = angle * Math.PI / 180.0;
					double num2 = (crop ? ((double)source.Width) : (Math.Abs((double)source.Width * Math.Cos(num)) + Math.Abs((double)source.Height * Math.Sin(num))));
					double num3 = (double)rect.Width / num2 / (crop ? PreviewGeometry.CropScale(source.Width, source.Height, angle) : 1.0);
					g.TranslateTransform(rect.X + rect.Width / 2f, rect.Y + rect.Height / 2f);
					g.RotateTransform((float)angle);
					g.ScaleTransform((float)num3, (float)num3);
					g.InterpolationMode = InterpolationMode.HighQualityBicubic;
					g.PixelOffsetMode = PixelOffsetMode.HighQuality;
					using (ImageAttributes imageAttributes = new ImageAttributes())
					{
						imageAttributes.SetWrapMode(WrapMode.TileFlipXY);
						g.DrawImage(source, new PointF[3]
						{
							new PointF((float)(-source.Width) / 2f, (float)(-source.Height) / 2f),
							new PointF((float)source.Width / 2f, (float)(-source.Height) / 2f),
							new PointF((float)(-source.Width) / 2f, (float)source.Height / 2f)
						}, new RectangleF(0f, 0f, source.Width, source.Height), GraphicsUnit.Pixel, imageAttributes);
					}
				}
				finally
				{
					g.Restore(gstate2);
				}
				if (grid)
				{
					using (Pen pen = new Pen(Color.FromArgb(150, 20, 35, 50), 2.5f))
					{
						using (Pen pen2 = new Pen(Color.FromArgb(230, 255, 255, 255), 1f))
						{
							pen.DashPattern = new float[2] { 2f, 2f };
							pen2.DashPattern = new float[2] { 5f, 5f };
							for (int k = 1; k < Math.Min(50, columns); k++)
							{
								float num4 = rect.Left + rect.Width * (float)k / (float)columns;
								g.DrawLine(pen, num4, rect.Top, num4, rect.Bottom);
								g.DrawLine(pen2, num4, rect.Top, num4, rect.Bottom);
							}
							for (int l = 1; l < Math.Min(50, rows); l++)
							{
								float num5 = rect.Top + rect.Height * (float)l / (float)rows;
								g.DrawLine(pen, rect.Left, num5, rect.Right, num5);
								g.DrawLine(pen2, rect.Left, num5, rect.Right, num5);
							}
						}
					}
				}
			}
			finally
			{
				g.Restore(gstate);
			}
			using (Pen pen3 = new Pen(Color.FromArgb(120, 137, 155)))
			{
				g.DrawRectangle(pen3, rect.X, rect.Y, rect.Width, rect.Height);
			}
		}
	}
	public class PhotoSnapshot
	{
		public object Shape;

		public int Index;

		public int Id;

		public double Left;

		public double Top;

		public double Width;

		public double Height;

		public double Rotation;

		public string Name;

		public string AlternativeText;

		public PhotoBox Box()
		{
			PhotoBox photoBox = new PhotoBox();
			photoBox.Id = Id;
			photoBox.Left = Left;
			photoBox.Top = Top;
			photoBox.Width = Width;
			photoBox.Height = Height;
			return photoBox;
		}
	}
	public class SelectionSnapshot
	{
		public object Slide;

		public object Presentation;

		public double SlideWidth;

		public double SlideHeight;

		public List<PhotoSnapshot> Photos = new List<PhotoSnapshot>();

		public List<PhotoBox> Boxes()
		{
			return Photos.ConvertAll((PhotoSnapshot p) => p.Box());
		}
	}
	public partial class PowerPointHost
	{
		private readonly dynamic app;

		public IntPtr WindowHandle
		{
			get
			{
				try
				{
					return new IntPtr((int)app.HWND);
				}
				catch
				{
					return IntPtr.Zero;
				}
			}
		}

		public PowerPointHost(object application)
		{
			if (application == null)
			{
				throw new InvalidOperationException("PowerPoint에 연결되지 않았습니다.");
			}
			app = application;
		}

		public SelectionSnapshot ReadSelection()
		{
			try
			{
				SelectionSnapshot selection = ReadSelectedPhotosOrNull();
				if (selection == null)
					throw new InvalidOperationException("슬라이드에서 편집할 사진을 먼저 선택하세요. 여러 장은 Ctrl 키를 누르며 선택할 수 있습니다.");
				if (selection.Photos.Count == 0)
					throw new InvalidOperationException("선택한 항목 중 사진을 먼저 선택해 주세요.");
				return selection;
			}
			catch (InvalidOperationException)
			{
				throw;
			}
			catch (Exception innerException)
			{
				throw new InvalidOperationException("일반 편집 화면에서 사진을 선택한 뒤 다시 시도하세요.", innerException);
			}
		}

		// A mixed shape range is normal in PowerPoint. Layout tools use only
		// pictures from that range and intentionally leave every other shape alone.
		internal SelectionSnapshot ReadSelectedPhotosOrNull()
		{
			try
			{
				dynamic window = app.ActiveWindow;
				dynamic selection = window.Selection;
				if ((int)selection.Type != 2) return null;
				if ((bool)selection.HasChildShapeRange)
					throw new InvalidOperationException("그룹 내부 사진은 그룹을 해제한 뒤 선택하세요.");
				dynamic range = selection.ShapeRange;
				dynamic presentation = app.ActivePresentation;
				SelectionSnapshot snapshot = new SelectionSnapshot
				{
					Slide = window.View.Slide,
					Presentation = presentation,
					SlideWidth = (double)presentation.PageSetup.SlideWidth,
					SlideHeight = (double)presentation.PageSetup.SlideHeight
				};
				for (int i = 1; i <= (int)range.Count; i++)
				{
					dynamic shape = range.Item(i);
					CollectSelectedPhotos(shape, snapshot.Photos);
				}
				if (snapshot.Photos.Count > 50)
					throw new InvalidOperationException("한 번에 사진 50장까지 선택할 수 있습니다.");
				return snapshot;
			}
			catch (InvalidOperationException) { throw; }
			catch (Exception innerException)
			{
				throw new InvalidOperationException("일반 편집 화면에서 사진을 선택한 뒤 다시 시도하세요.", innerException);
			}
		}

		private static bool IsSelectedPhoto(dynamic shape)
		{
			int type = (int)shape.Type;
			if (type == 13 || type == 11) return true;
			if (type != 14) return false;
			try
			{
				int contained = (int)shape.PlaceholderFormat.ContainedType;
				return contained == 13 || contained == 11;
			}
			catch { return false; }
		}

		private static void CollectSelectedPhotos(dynamic shape, List<PhotoSnapshot> photos)
		{
			if ((int)shape.Type == 6)
			{
				for (int i = 1; i <= (int)shape.GroupItems.Count; i++)
					CollectSelectedPhotos(shape.GroupItems.Item(i), photos);
				return;
			}
			if (IsSelectedPhoto(shape)) photos.Add(SnapshotPhoto(shape));
		}

		private static PhotoSnapshot SnapshotPhoto(dynamic shape)
		{
			return new PhotoSnapshot
			{
				Shape = shape,
				Id = (int)shape.Id,
				Index = (int)shape.ZOrderPosition,
				Left = (double)shape.Left,
				Top = (double)shape.Top,
				Width = (double)shape.Width,
				Height = (double)shape.Height,
				Rotation = (double)shape.Rotation,
				Name = (string)shape.Name,
				AlternativeText = (string)shape.AlternativeText
			};
		}

		public void ApplyLayout(SelectionSnapshot selection, List<PhotoBox> plan)
		{
			VerifyUnchanged(selection);
			foreach (PhotoSnapshot photo in selection.Photos)
			{
				if (Math.Abs(Normalize(photo.Rotation)) > 0.001)
				{
					throw new InvalidOperationException("회전된 사진은 먼저 사진 회전에서 반듯하게 자르기를 적용한 뒤 배열하세요.");
				}
			}
			app.StartNewUndoEntry();
			try
			{
				foreach (PhotoBox box in plan)
				{
					List<PhotoSnapshot> photos = selection.Photos;
					Predicate<PhotoSnapshot> match = (PhotoSnapshot p) => p.Id == box.Id;
					PhotoSnapshot photoSnapshot = photos.Find(match);
					if (photoSnapshot == null)
					{
						throw new InvalidOperationException("선택한 사진이 변경되었습니다. 창을 닫고 다시 선택하세요.");
					}
					if (!SetAttachedLabelGroupBox(photoSnapshot, box)) SetBox(photoSnapshot.Shape, box);
				}
			}
			catch
			{
				foreach (PhotoSnapshot photo2 in selection.Photos)
				{
					try
					{
						if (!SetAttachedLabelGroupBox(photo2, photo2.Box())) SetBox(photo2.Shape, photo2.Box());
					}
					catch
					{
					}
				}
				throw;
			}
		}

		private static void SetBox(dynamic shape, PhotoBox box)
		{
			int num = (int)shape.LockAspectRatio;
			try
			{
				shape.LockAspectRatio = 0;
				shape.Width = (float)box.Width;
				shape.Height = (float)box.Height;
				shape.Left = (float)box.Left;
				shape.Top = (float)box.Top;
			}
			finally
			{
				shape.LockAspectRatio = num;
			}
		}

		private static double Normalize(double angle)
		{
			return ((angle + 180.0) % 360.0 + 360.0) % 360.0 - 180.0;
		}

		private void VerifyUnchanged(SelectionSnapshot selection)
		{
			foreach (PhotoSnapshot photo in selection.Photos)
			{
				dynamic shape = photo.Shape;
				if ((int)shape.Id != photo.Id || Math.Abs((double)shape.Left - photo.Left) > 0.01 || Math.Abs((double)shape.Top - photo.Top) > 0.01 || Math.Abs((double)shape.Width - photo.Width) > 0.01 || Math.Abs((double)shape.Height - photo.Height) > 0.01 || Math.Abs(Normalize((double)shape.Rotation - photo.Rotation)) > 0.01)
				{
					throw new InvalidOperationException("사진이 변경되었습니다. 창을 닫고 다시 선택하세요.");
				}
			}
		}

		private void ExportPhoto(dynamic shape, SelectionSnapshot selection, string path)
		{
			double num = (double)shape.Width * (double)shape.Height * Math.Pow(4.166666666666667, 2.0);
			if (num > 40000000.0)
			{
				throw new InvalidOperationException("사진이 너무 큽니다. 슬라이드에서 크기를 줄인 뒤 다시 시도하세요.");
			}
			shape.Export(path, 2, (int)Math.Round(selection.SlideWidth * 300.0 / 96.0), (int)Math.Round(selection.SlideHeight * 300.0 / 96.0), 1);
			if (!File.Exists(path))
			{
				throw new InvalidOperationException("PowerPoint에서 사진을 읽지 못했습니다.");
			}
		}

		internal virtual ImagePreviewSession CreatePreviewSession(SelectionSnapshot selection)
		{
			VerifyUnchanged(selection);
			string text = Engine.NewJobDirectory();
			dynamic val = null;
			try
			{
				string text2 = Path.Combine(text, "preview.pptx");
				((dynamic)selection.Presentation).SaveCopyAs(text2, 24);
				val = app.Presentations.Open(text2, -1, 0, 0);
				dynamic val2 = val.Slides.Item((int)((dynamic)selection.Slide).SlideIndex);
				dynamic val3 = val2.Shapes.Item(selection.Photos[0].Index);
				val3.Rotation = 0f;
				ExportPhoto(val3, selection, Path.Combine(text, "input.png"));
				val.Saved = -1;
				val.Close();
				val = null;
				File.Delete(text2);
				return new ImagePreviewSession(text);
			}
			catch
			{
				Engine.CleanJob(text);
				throw;
			}
			finally
			{
				if (val != null)
				{
					try
					{
						val.Saved = -1;
						val.Close();
						Engine.CleanJob(text);
					}
					catch
					{
					}
				}
			}
		}

		public async Task<Image> PreviewAsync(SelectionSnapshot selection, bool removeBackground, bool straighten, double angle, CancellationToken token)
		{
			VerifyOptions(removeBackground, straighten, angle);
			Engine.ValidateReady();
			VerifyUnchanged(selection);
			string job = Engine.NewJobDirectory();
			dynamic scratchPresentation = null;
			try
			{
				string presentationPath = Path.Combine(job, "preview.pptx");
				((dynamic)selection.Presentation).SaveCopyAs(presentationPath, 24);
				scratchPresentation = app.Presentations.Open(presentationPath, -1, 0, 0);
				dynamic scratch = scratchPresentation.Slides.Item((int)((dynamic)selection.Slide).SlideIndex);
				PhotoSnapshot photo = selection.Photos[0];
				dynamic shape = scratch.Shapes.Item(photo.Index);
				shape.Rotation = 0f;
				string input = Path.Combine(job, "input.png");
				string output = Path.Combine(job, "output.png");
				ExportPhoto(shape, selection, input);
				scratchPresentation.Saved = -1;
				scratchPresentation.Close();
				scratchPresentation = null;
				await Engine.RunAsync(Operation(removeBackground, straighten), new List<EngineItem>
				{
					new EngineItem
					{
						input = input,
						output = output,
						angle = (straighten ? Normalize(photo.Rotation + angle) : 0.0)
					}
				}, job, token);
				using (Image image = Image.FromFile(output))
				{
					if (straighten || Math.Abs(Normalize(photo.Rotation)) < 0.001)
					{
						return new Bitmap(image);
					}
					return RotatePreview(image, photo.Rotation);
				}
			}
			finally
			{
				if (scratchPresentation != null)
				{
					try
					{
						scratchPresentation.Saved = -1;
						scratchPresentation.Close();
					}
					catch
					{
					}
				}
				Engine.CleanJob(job);
			}
		}

		private static Image RotatePreview(Image source, double angle)
		{
			double num = angle * Math.PI / 180.0;
			int num2 = (int)Math.Ceiling(Math.Abs((double)source.Width * Math.Cos(num)) + Math.Abs((double)source.Height * Math.Sin(num)));
			int num3 = (int)Math.Ceiling(Math.Abs((double)source.Width * Math.Sin(num)) + Math.Abs((double)source.Height * Math.Cos(num)));
			Bitmap bitmap = new Bitmap(num2, num3);
			using (Graphics graphics = Graphics.FromImage(bitmap))
			{
				graphics.TranslateTransform((float)num2 / 2f, (float)num3 / 2f);
				graphics.RotateTransform((float)angle);
				graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
				graphics.DrawImage(source, (float)(-source.Width) / 2f, (float)(-source.Height) / 2f, source.Width, source.Height);
				return bitmap;
			}
		}

		public async Task ApplyImagesAsync(SelectionSnapshot selection, bool removeBackground, bool straighten, double angle, CancellationToken token)
		{
			VerifyOptions(removeBackground, straighten, angle);
			Engine.ValidateReady();
			VerifyUnchanged(selection);
			string job = Engine.NewJobDirectory();
			dynamic copy = null;
			bool committed = false;
			try
			{
				app.StartNewUndoEntry();
				copy = ((dynamic)selection.Slide).Duplicate().Item(1);
				List<EngineItem> items = new List<EngineItem>();
				for (int i = 0; i < selection.Photos.Count; i++)
				{
					token.ThrowIfCancellationRequested();
					PhotoSnapshot photoSnapshot = selection.Photos[i];
					dynamic val = copy.Shapes.Item(photoSnapshot.Index);
					val.Rotation = 0f;
					string text = Path.Combine(job, "in-" + i + ".png");
					string output = Path.Combine(job, "out-" + i + ".png");
					ExportPhoto(val, selection, text);
					items.Add(new EngineItem
					{
						input = text,
						output = output,
						angle = (straighten ? Normalize(photoSnapshot.Rotation + angle) : 0.0)
					});
				}
				await Engine.RunAsync(Operation(removeBackground, straighten), items, job, token);
				VerifyUnchanged(selection);
				token.ThrowIfCancellationRequested();
				for (int j = 0; j < selection.Photos.Count; j++)
				{
					PhotoSnapshot photoSnapshot2 = selection.Photos[j];
					dynamic val2 = copy.Shapes.Item(photoSnapshot2.Index);
					dynamic val3 = copy.Shapes.AddPicture(items[j].output, 0, -1, (float)photoSnapshot2.Left, (float)photoSnapshot2.Top, (float)photoSnapshot2.Width, (float)photoSnapshot2.Height);
					val3.Rotation = (straighten ? 0f : ((float)photoSnapshot2.Rotation));
					val3.AlternativeText = photoSnapshot2.AlternativeText;
					val3.Name = "LabPhoto_" + photoSnapshot2.Id;
					while ((int)val3.ZOrderPosition > photoSnapshot2.Index)
					{
						val3.ZOrder(3);
					}
					val2.Delete();
					val3.Name = photoSnapshot2.Name;
				}
				committed = true;
				try
				{
					app.ActiveWindow.View.GotoSlide((int)copy.SlideIndex);
				}
				catch
				{
				}
			}
			finally
			{
				if (!committed && copy != null)
				{
					try
					{
						copy.Delete();
					}
					catch
					{
					}
				}
				Engine.CleanJob(job);
			}
		}

		private static string Operation(bool background, bool straighten)
		{
			if (!background)
			{
				return "straighten";
			}
			if (!straighten)
			{
				return "remove_background";
			}
			return "both";
		}

		private static void VerifyOptions(bool background, bool straighten, double angle)
		{
			if (!background && !straighten)
			{
				throw new InvalidOperationException("적용할 사진 회전 기능을 하나 이상 선택하세요.");
			}
			if (double.IsNaN(angle) || double.IsInfinity(angle) || Math.Abs(angle) > 45.0)
			{
				throw new InvalidOperationException("추가 회전 각도는 -45°에서 45° 사이로 입력하세요.");
			}
		}
	}
}
