using System;
using System.Collections.Generic;

using FastExcel;
using Bau.Libraries.LibExcelFiles.Content;

namespace Bau.Libraries.LibExcelFiles
{
	/// <summary>
	///		Manager para las librerías de excel
	/// </summary>
    public class ExcelManager : IDisposable
    {
		// Variables privadas
		private FastExcel.FastExcel _excel;
		private Worksheet _workSheet;

		/// <summary>
		///		Carga un archivo Excel
		/// </summary>
		public void Load(string fileName, bool readOnly)
		{
			_excel = new FastExcel.FastExcel(new System.IO.FileInfo(fileName), readOnly);
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
			foreach (Worksheet sheet in _excel.Worksheets)
				yield return sheet.Name;
		}

		/// <summary>
		///		Cambia la hoja
		/// </summary>
		public void SetWorksheet(string sheet)
		{
			_workSheet = _excel.Read(sheet);
		}

		/// <summary>
		///		Cambia la hoja
		/// </summary>
		public void SetWorksheet(int sheet)
		{
			_workSheet = _excel.Read(sheet);
		}

		/// <summary>
		///		Obtiene las filas
		/// </summary>
		public IEnumerable<RowModel> GetRows()
		{
			foreach (Row row in _workSheet.Rows)
				yield return ConvertRow(row);
		}

		/// <summary>
		///		Convierte una fila
		/// </summary>
		private RowModel ConvertRow(Row row)
		{
			RowModel result = new RowModel(row.RowNumber);

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
}
