using System.Data;

using FastExcel;

namespace Bau.Libraries.LibExcelFiles.Data;

/// <summary>
///		Generador de un Excel a partir de un DataTable
/// </summary>
public class ExcelDataTableWriter
{
	/// <summary>
	///		Escribe un DataTable en un archivo Excel
	/// </summary>
	public void Write(string template, string fileName, string sheetName, DataTable data)
	{
		Worksheet sheet = new();
		List<Row> rows = [];

			// Añade la cabecera de la tabla de datos
			rows.Add(CreateHeaderRow(data));
			// Añade las filas
			rows.AddRange(CreateDataRows(data));
			// Asigna las filas a la hoja
			sheet.Rows = rows;
			// Crea el archivo
			if (!File.Exists(fileName))
				try
				{
					File.Delete(fileName);
				}
				catch {}
			// Escribe los datos
			using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(new FileInfo(template), new FileInfo(fileName)))
			{
				fastExcel.Worksheets[0].Name = sheetName;
				fastExcel.Write(sheet, sheetName);
			}
	}

	/// <summary>
	///		Crea una fila de cabecera para el Excel con las columnas de la tabla
	/// </summary>
	private Row CreateHeaderRow(DataTable data)
	{
		List<Cell> cells = [];
		int index = 1;

			// Añade los nombres de los campos
			foreach (DataColumn column in data.Columns)
				cells.Add(new Cell(index++, column.ColumnName));
			// Devuelve la fila creada
			return new Row(1, cells);
	}

	/// <summary>
	///		Crea los datos para las filas del <see cref="DataTable"/>
	/// </summary>
	private List<Row> CreateDataRows(DataTable data)
	{
		List<Row> rows = [];

			// Añade las filas
			foreach (DataRow row in data.Rows)
			{
				List<Cell> cells = [];

					// Crea las celdas
					for (int column = 0; column < data.Columns.Count; column++)
						cells.Add(new Cell(column + 1, row[column]));
					// Añade las celdas a la nueva fila
					rows.Add(new Row(rows.Count + 1, cells));
			}
			// Devuelve la colección de filas
			return rows;
	}
}