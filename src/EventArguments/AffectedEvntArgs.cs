using System;

namespace Bau.Libraries.LibExcelFiles.EventArguments
{
	/// <summary>
	///		Argumento del evento de lectura / escritura sobre un archivo Excel
	/// </summary>
	public class AffectedEvntArgs : EventArgs
	{
		public AffectedEvntArgs(long records)
		{
			Records = records;
		}

		/// <summary>
		///		Número de registros
		/// </summary>
		public long Records { get; }
	}
}
