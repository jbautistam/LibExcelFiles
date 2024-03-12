namespace Bau.Libraries.LibExcelFiles.Content;

/// <summary>
///		Clase con los datos de una celda
/// </summary>
public class CellModel
{
	/// <summary>
	///		Fila de la celda
	/// </summary>
	public int Row { get; set; }

	/// <summary>
	///		Columna de la celda
	/// </summary>
	public int Column { get; set; }

	/// <summary>
	///		Valor de la celda
	/// </summary>
	public object Value { get; set; }
}