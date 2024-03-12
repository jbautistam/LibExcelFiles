using FastExcel;
using Bau.Libraries.LibExcelFiles.Content;

namespace Bau.Libraries.LibExcelFiles;

/// <summary>
///		Manager para las librerías de excel
/// </summary>
public class ExcelManager : IDisposable
{
	// Variables privadas
	private FastExcel.FastExcel? _excel;
	private Worksheet? _workSheet;

	/// <summary>
	///		Carga un archivo Excel
	/// </summary>
	public void Load(string fileName, bool readOnly)
	{
		_excel = new FastExcel.FastExcel(new FileInfo(fileName), readOnly);
	}

	/// <summary>
	///		Carga un archivo Excel
	/// </summary>
	public void Load(Stream stream)
	{
		_excel = new FastExcel.FastExcel(stream);
	}

	/// <summary>
	///		Crea una hoja
	/// </summary>
	public void CreateWorksheet(string name)
	{
		_workSheet = new Worksheet(_excel);
		_workSheet.Name = name;
	}

	/// <summary>
	///		Obtiene los nombres de las hojas
	/// </summary>
	public IEnumerable<string> GetWorkSheets()
	{
		if (_excel is not null)
			foreach (Worksheet sheet in _excel.Worksheets)
				yield return sheet.Name;
	}

	/// <summary>
	///		Cambia la hoja
	/// </summary>
	public void SetWorksheet(string sheet)
	{
		_workSheet = _excel?.Read(sheet);
	}

	/// <summary>
	///		Cambia la hoja
	/// </summary>
	public void SetWorksheet(int sheet)
	{
		_workSheet = _excel?.Read(sheet);
	}

	/// <summary>
	///		Obtiene una fila
	/// </summary>
	public RowModel? GetRow(int rowIndex)
	{
		// Busca la fila
		if (_workSheet is not null)
			foreach (Row row in _workSheet.Rows)
				if (row.RowNumber == rowIndex)
					return ConvertRow(row);
		// Si ha llegado hasta aquí es porque no ha encontrado nada
		return null;
	}

	/// <summary>
	///		Obtiene las filas
	/// </summary>
	public IEnumerable<RowModel> GetRows()
	{
		if (_workSheet is not null)
			foreach (Row row in _workSheet.Rows)
				yield return ConvertRow(row);
	}

	/// <summary>
	///		Convierte una fila
	/// </summary>
	private RowModel ConvertRow(Row row)
	{
		RowModel result = new(row.RowNumber);

			// Añade las celdas
			foreach (Cell cell in row.Cells)
				result.Cells.Add(new CellModel
										{
											Column = cell.ColumnNumber,
											Row = cell.RowNumber,
											Value = cell.Value
										}
								);
			// Devuelve la fila
			return result;
	}

	/// <summary>
	///		Cierra el archivo Excel
	/// </summary>
	public void Close()
	{
		if (_excel != null)
		{
			_excel.Dispose();
			_excel = null;
		}
	}

	/// <summary>
	///		Libera la memoria
	/// </summary>
	protected virtual void Dispose(bool disposing)
	{
		if (!IsDisposed)
		{
			// Libera la memoria
			if (disposing)
				Close();
			// Indica que se ha liberado la memoria
			IsDisposed = true;
		}
	}

	/// <summary>
	///		Libera la memoria
	/// </summary>
	public void Dispose()
	{
		Dispose(true);
	}

	/// <summary>
	///		Indica si se ha liberado la memoria
	/// </summary>
	public bool IsDisposed { get; private set; }
}
