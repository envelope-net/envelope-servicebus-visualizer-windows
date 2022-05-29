using Envelope.ServiceBus.Orchestrations.Graphing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Envelope.ServiceBus.Visualizer.Windows;

/// <summary>
/// Graph image generator
/// </summary>
public static class OrchestrationImageGenerator
{
	internal const string DOT_EXE = @".\DLL\dot.exe";

	internal const string LIB_GVC = @".\DLL\gvc.dll";
	internal const string LIB_GRAPH = @".\DLL\cgraph.dll";
	internal const int SUCCESS = 0;

	public const string bmp = "bmp";
	public const string eps = "eps";
	public const string gif = "gif";
	public const string jpg = "jpg";
	public const string pdf = "pdf";
	public const string png = "png";
	public const string svg = "svg";
	public const string tif = "tif";
	public const string vml = "vml";
	public const string webp = "webp";

	#region pInvoke

	/// <summary>
	/// Creates a new Graphviz context.
	/// </summary>
	[DllImport(LIB_GVC, CallingConvention = CallingConvention.Cdecl)]
	internal static extern IntPtr gvContext();

	/// <summary>
	/// Releases a context's resources.
	/// </summary>
	[DllImport(LIB_GVC, CallingConvention = CallingConvention.Cdecl)]
	internal static extern int gvFreeContext(IntPtr gvc);

	/// <summary>
	/// Reads a graph from a string.
	/// </summary>
	[DllImport(LIB_GRAPH, CallingConvention = CallingConvention.Cdecl)]
	internal static extern IntPtr agmemread(string data);

	/// <summary>
	/// Releases the resources used by a graph.
	/// </summary>
	[DllImport(LIB_GRAPH, CallingConvention = CallingConvention.Cdecl)]
	internal static extern void agclose(IntPtr g);

	/// <summary>
	/// Applies a layout to a graph using the given engine.
	/// </summary>
	[DllImport(LIB_GVC, CallingConvention = CallingConvention.Cdecl)]
	internal static extern int gvLayout(IntPtr gvc, IntPtr g, string engine);

	/// <summary>
	/// Releases the resources used by a layout.
	/// </summary>
	[DllImport(LIB_GVC, CallingConvention = CallingConvention.Cdecl)]
	internal static extern int gvFreeLayout(IntPtr gvc, IntPtr g);

	/// <summary>
	/// Renders a graph to a file.
	/// </summary>
	[DllImport(LIB_GVC, CallingConvention = CallingConvention.Cdecl)]
	internal static extern int gvRenderFilename(IntPtr gvc, IntPtr g,
		  string format, string fileName);

	/// <summary>
	/// Renders a graph in memory.
	/// </summary>
	[DllImport(LIB_GVC, CallingConvention = CallingConvention.Cdecl)]
	internal static extern int gvRenderData(IntPtr gvc, IntPtr g,
		 string format, out IntPtr result, out int length);

	/// <summary>
	/// Release render resources.
	/// </summary>
	[DllImport(LIB_GVC, CallingConvention = CallingConvention.Cdecl)]
	internal static extern int gvFreeRenderData(IntPtr result);

	#endregion pInvoke

	/// <summary>
	/// Create graph image by calling dot.exe save temp files to disk, read and delete files
	/// </summary>
	public static MemoryStream CreateImageFile(string dotGraph, string outputFormat, string absoluteTempFolderPath)
	{
		if (!OperatingSystem.IsWindows())
			throw new NotSupportedException("Only Microsoft Windows is supported");

		if (string.IsNullOrWhiteSpace(dotGraph))
			throw new ArgumentNullException(nameof(dotGraph));

		if (string.IsNullOrWhiteSpace(outputFormat))
			throw new ArgumentNullException(nameof(outputFormat));

		if (string.IsNullOrWhiteSpace(absoluteTempFolderPath))
			throw new ArgumentNullException(nameof(absoluteTempFolderPath));


		Directory.CreateDirectory(absoluteTempFolderPath);
		var inputFileName = $"{Guid.NewGuid():D}.dot";
		var inputFilePath = Path.Combine(absoluteTempFolderPath, inputFileName);
		File.WriteAllText(inputFilePath, dotGraph, new UTF8Encoding(false));

		var outputFileName = $"{Guid.NewGuid():D}.{outputFormat}";
		var outputFilePath = Path.Combine(absoluteTempFolderPath, outputFileName);

		var process = new Process();

		// Stop the process from opening a new window
		process.StartInfo.RedirectStandardOutput = true;
		process.StartInfo.UseShellExecute = false;
		process.StartInfo.CreateNoWindow = true;

		// Setup executable and parameters
		process.StartInfo.FileName = DOT_EXE;
		process.StartInfo.Arguments = string.Format($"{inputFilePath} -T{outputFormat} -o {outputFilePath}");

		// Go
		process.Start();
		// and wait dot.exe to complete and exit
		process.WaitForExit();

		var ms = new MemoryStream();
		using (var fileStream = File.Open(outputFilePath, FileMode.Open))
		{
			//var image = Image.FromStream(bmpStream);
			//bitmap = new Bitmap(image);

			fileStream.CopyTo(ms);
			ms.Flush();
		}

		File.Delete(inputFilePath);
		File.Delete(outputFilePath);

		ms.Seek(0, SeekOrigin.Begin);
		return ms;
	}

	/// <summary>
	/// Create graph image by PInvoke on gvc.dll and cgraph.dll
	/// </summary>
	public static MemoryStream CreateImage(string dotGraph, string outputFormat)
	{
		// Create a Graphviz context
		IntPtr gvc = gvContext();
		if (gvc == IntPtr.Zero)
			throw new InvalidOperationException("Failed to create Graphviz context.");

		// Load the DOT data into a graph
		IntPtr g = agmemread(dotGraph);
		if (g == IntPtr.Zero)
			throw new InvalidOperationException("Failed to create graph from source. Check for syntax errors.");

		// Apply a layout
		if (gvLayout(gvc, g, "dot") != SUCCESS)
			throw new InvalidOperationException("Layout failed.");

		// Render the graph
		if (gvRenderData(gvc, g, outputFormat, out IntPtr result, out int length) != SUCCESS)
			throw new InvalidOperationException("Render failed.");

		// Create an array to hold the rendered graph
		var bytes = new byte[length];

		// Copy the image from the IntPtr
		Marshal.Copy(result, bytes, 0, length);

		// Free up the resources
		_ = gvFreeLayout(gvc, g);
		agclose(g);
		_ = gvFreeContext(gvc);
		_ = gvFreeRenderData(result);

		var ms = new MemoryStream(bytes);
		ms.Seek(0, SeekOrigin.Begin);
		return ms;
	}

	public static MemoryStream CreateImageFile(
		IOrchestrationGraph orchestrationGraph,
		string outputFormat,
		string absoluteTempFolderPath,
		IEnumerable<ExecutionPointerColor>? executionPointers = null)
	{
		if (orchestrationGraph == null)
			throw new ArgumentNullException(nameof(orchestrationGraph));

		var generator = new OrchestrationGraphvizGenerator(orchestrationGraph, executionPointers);
		var dotGraph = generator.CreateDotGraph();
		return CreateImageFile(dotGraph, outputFormat, absoluteTempFolderPath);
	}

	public static MemoryStream CreateImage(
		IOrchestrationGraph orchestrationGraph,
		string outputFormat,
		IEnumerable<ExecutionPointerColor>? executionPointers = null)
	{
		if (orchestrationGraph == null)
			throw new ArgumentNullException(nameof(orchestrationGraph));

		var generator = new OrchestrationGraphvizGenerator(orchestrationGraph, executionPointers);
		var dotGraph = generator.CreateDotGraph();
		return CreateImage(dotGraph, outputFormat);
	}
}
