using System;

namespace Bau.Libraries.LibExcelFiles.Content
{
	/// <summary>
	///		Clase con los datos de una fila de excel
	/// </summary>
    public class RowModel
    {
		public RowModel(int rowNumber)
		{
			RowNumber = rowNumber;
		}

		/// <summary>
		///		Obtiene el número máximo de columna
		/// </summary>
		public int GetMaxColumn()
		{
			int maximum = 0;

				// Recorre las celdas buscando el máximo valor de columna
				foreach (CellModel cell in Cells)
					if (cell.Column > maximum)
						maximum = cell.Column;
				// Devuelve el índice máximo
				return maximum;
		}

		/// <summary>
		///		Obtiene la celda de una columna
		/// </summary>
		public CellModel GetCell(int column)
		{
			// Busca la celda
			foreach (CellModel cell in Cells)
				if (cell.Column == column)
					return cell;
			// Si ha llegado hasta aquí es porque no ha encontrado nada
			return null;
		}

		/// <summary>
		///		Número de fila
		/// </summary>
		public int RowNumber { get; }

		/// <summary>
		///		Celdas de la fila
		/// </summary>
		public System.Collections.Generic.List<CellModel> Cells { get; } = new System.Collections.Generic.List<CellModel>();
    }
}
