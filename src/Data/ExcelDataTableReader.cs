using System;
using System.Data;

using Bau.Libraries.LibExcelFiles.Content;

namespace Bau.Libraries.LibExcelFiles.Data
{
	/// <summary>
	///		Lector de Excel para obtención de un DataTable
	/// </summary>
    public class ExcelDataTableReader
    {
		/// <summary>
		///		Carga los datos del Excel en un <see cref="DataTable"/>
		/// </summary>
		public DataTable LoadFile(string fileName, int sheetIndex, int startRow, long rows, bool withHeader)
		{
			if (sheetIndex < 1)
				throw new ArgumentException("Sheet index must be greater than 0");
			else
				using (ExcelManager manager = Load(fileName))
				{
					// Ajusta la hoja
					manager.SetWorksheet(sheetIndex);
					// Crea el dataTable
					return Create(manager, startRow, rows, withHeader);
				}
		}

		/// <summary>
		///		Carga los datos del Excel en un <see cref="DataTable"/>
		/// </summary>
		public DataTable LoadFile(string fileName, string sheet, int startRow, long rows, bool withHeader)
		{
			using (ExcelManager manager = Load(fileName))
			{
				// Ajusta la hoja
				manager.SetWorksheet(sheet);
				// Crea el dataTable
				return Create(manager, startRow, rows, withHeader);
			}
		}

		/// <summary>
		///		Obtiene el número de filas
		/// </summary>
		public long CountRows(string fileName, int sheetIndex, bool withHeader)
		{
			if (sheetIndex < 1)
				throw new ArgumentException("Sheet index must be greater than 0");
			else
				using (ExcelManager manager = Load(fileName))
				{
					// Ajusta la hoja
					manager.SetWorksheet(sheetIndex);
					// Devuelve el número de filas
					return CountRows(manager, withHeader);
				}
		}

		/// <summary>
		///		Obtiene el número de filas
		/// </summary>
		public long CountRows(string fileName, string sheetName, bool withHeader)
		{
			using (ExcelManager manager = Load(fileName))
			{
				// Ajusta la hoja
				manager.SetWorksheet(sheetName);
				// Devuelve el número de filas
				return CountRows(manager, withHeader);
			}
		}

		/// <summary>
		///		Cuenta las filas de la hoja activa
		/// </summary>
		private long CountRows(ExcelManager manager, bool withHeader)
		{
			long rows = 0;

				// Cuenta los datos
				foreach (RowModel row in manager.GetRows())
					rows++;
				// Quita la fila de cabecera
				if (withHeader)
					rows--;
				// Devuelve el número de filas
				return rows;
		}

		/// <summary>
		///		Carga el archivo
		/// </summary>
		private ExcelManager Load(string fileName)
		{
			ExcelManager manager = new ExcelManager();

				// Carga el archivo
				manager.Load(fileName, true);
				// Devuelve el manager con el archivo cargado
				return manager;
		}

		/// <summary>
		///		Crea la dataTabla con los datos del excel
		/// </summary>
		private DataTable Create(ExcelManager manager, int startRow, long rows, bool withHeader)
		{
			DataTable table = new DataTable();
			long rowsRead = 0;

				// Recorre las filas
				foreach (RowModel row in manager.GetRows())
					if (row.RowNumber >= startRow)
					{
						bool skip = false;

							// Crea las cabeceras del dataTable
							if (table.Columns.Count == 0)
							{
								if (!withHeader)
									CreateColumns(table, row.GetMaxColumn());
								else
								{
									CreateColumns(table, row);
									skip = true;
								}
							}
							// Crea la fila de datos y sale del bucle si es necesario
							if (!skip)
							{
								// Crea la fila
								CreateRow(table, row);
								// Sale del bucle si es necesario
								if (rows != 0 && ++rowsRead > rows)
									break;
							}
					}
				// Devuelve la tabla de datos
				return table;
		}

		/// <summary>
		///		Crea las columnas en el <see cref="DataTable"/> a partir de los datos de una fila de Excel
		/// </summary>
		private void CreateColumns(DataTable table, RowModel row)
		{
			for (int column = 1; column <= row.GetMaxColumn(); column++)
			{
				CellModel cell = row.GetCell(column);

					if (cell?.Value == null)
						table.Columns.Add($"Field {column}");
					else
						table.Columns.Add(cell.Value.ToString());
			}
		}

		/// <summary>
		///		Crea una serie de columnas en el <see cref="DataTable"/>
		/// </summary>
		private void CreateColumns(DataTable table, int maxColumns)
		{
			for (int column = 1; column < maxColumns; column++)
				table.Columns.Add($"Field {column}");
		}

		/// <summary>
		///		Crea una fila sobre el <see cref="DataTable"/> con los datos leídos de Excel
		/// </summary>
		private void CreateRow(DataTable table, RowModel excelRow)
		{
			DataRow row = table.NewRow();

				// Añade las celdas
				for (int column = 1; column <= table.Columns.Count; column++)
				{
					CellModel cell = excelRow.GetCell(column);

						row[column - 1] = cell?.Value;
				}
				// Añade la fila a la tabla
				table.Rows.Add(row);
		}
    }
}
