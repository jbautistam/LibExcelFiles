using System.Data;

using FastExcel;

namespace Bau.Libraries.LibExcelFiles.Data;

/// <summary>
///		Generador de un Excel a partir de un <see cref="IDataReader"/>
/// </summary>
public class ExcelWriter
{
	public ExcelWriter(bool withHeader)
	{
		WithHeader = withHeader;
	}

	/// <summary>
	///		Escribe un <see cref="IDataReader"/> en un archivo Excel
	/// </summary>
	public void Write(Stream stream, int sheetIndex, IDataReader reader)
	{
		Worksheet sheet = new();
		List<Row> rows = [];

			// Añade la cabecera de la tabla de datos
			if (WithHeader)
				rows.Add(CreateHeaderRow(reader));
			// Añade las filas
			rows.AddRange(CreateDataRows(reader));
			// Asigna las filas a la hoja
			sheet.Rows = rows;
			// Escribe los datos
			using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(null, stream, false, false))
			{
				fastExcel.Write(sheet, sheetIndex);
			}
	}

	/// <summary>
	///		Crea una fila de cabecera para el Excel con las columnas de la tabla
	/// </summary>
	private Row CreateHeaderRow(IDataReader reader)
	{
		List<Cell> cells = [];

			// Añade los nombres de los campos
			for (int index = 0; index < reader.FieldCount; index++)
				cells.Add(new Cell(index + 1, reader.GetName(index)));
			// Devuelve la fila creada
			return new Row(1, cells);
	}

	/// <summary>
	///		Crea los datos para las filas del <see cref="DataTable"/>
	/// </summary>
	private List<Row> CreateDataRows(IDataReader reader)
	{
		List<Row> rows = [];

			// Añade las filas
			while (reader.Read())
			{
				List<Cell> cells = [];

					// Crea las celdas
					for (int column = 0; column < reader.FieldCount; column++)
						cells.Add(new Cell(column + 1, reader[column]));
					// Añade las celdas a la nueva fila
					rows.Add(new Row(rows.Count + 1, cells));
			}
			// Devuelve la colección de filas
			return rows;
	}

	/// <summary>
	///		Indica si tiene cabecera
	/// </summary>
	public bool WithHeader { get; }
}