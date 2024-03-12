using System.Data;

using Bau.Libraries.LibExcelFiles.Content;

namespace Bau.Libraries.LibExcelFiles.Data;

/// <summary>
///		Implementación de <see cref="IDataReader"/> para archivos Excel
/// </summary>
public class ExcelReader : IDataReader
{
	// Eventos públicos
	public event EventHandler<EventArguments.AffectedEvntArgs>? ReadBlock;
	// Variables privadas
	private ExcelManager _manager;
	private List<object> _recordsValues = [];
	private int _row;

	public ExcelReader(bool withHeader, int notifyAfter = 10_000)
	{
		_manager = new ExcelManager();
		WithHeader = withHeader;
		NotifyAfter = notifyAfter;
	}

	/// <summary>
	///		Abre el archivo
	/// </summary>
	public void Open(string fileName, int sheetIndex)
	{
		// Abre el archivo
		_manager.Load(fileName, true);
		// Inicializa el lector
		InitializeReader(sheetIndex);
	}

	/// <summary>
	///		Abre el datareader sobre el stream
	/// </summary>
	public void Open(Stream stream, int sheetIndex)
	{
		// Abre el archivo y cambia a la hoja adecuada
		_manager.Load(stream);
		// Inicializa el lector
		InitializeReader(sheetIndex);
	}

	/// <summary>
	///		Inicializa el lector
	/// </summary>
	private void InitializeReader(int sheetIndex)
	{
		// Cambia la hoja
		_manager.SetWorksheet(sheetIndex);
		// Carga las columnas
		LoadColumns();
		// Inicializa el número de fila actual
		_row = 1;
		if (WithHeader)
			_row = 2;
	}

	/// <summary>
	///		Carga las columnas del Excel
	/// </summary>
	private void LoadColumns()
	{
		RowModel rowHeader = _manager.GetRow(1);
		RowModel rowData = _manager.GetRow(2);

			// Añade los datos de la columna
			for (int index = 0; index < rowData.Cells.Count; index++)
			{
				string name = $"Column{index.ToString()}";

					// Si tenemos fila de cabecera, recogemos los datos de la cabecera
					if (WithHeader)
						name = rowHeader.Cells[index].Value?.ToString() ?? string.Empty;
					// Añadimos la columna
					Columns.Add((name, rowData.Cells[index].GetType()));
			}
	}

	/// <summary>
	///		Lee un registro
	/// </summary>
	public bool Read()
	{
		bool readed = false;
		RowModel line = _manager.GetRow(_row);

			// Interpreta los datos
			if (line != null)
			{
				// Interpreta la línea
				_recordsValues = new List<object>();
				foreach (CellModel cell in line.Cells)
					_recordsValues.Add(cell.Value);
				// Lanza el evento
				if (NotifyAfter > 0 && _row % NotifyAfter == 0)
					ReadBlock?.Invoke(this, new EventArguments.AffectedEvntArgs(_row));
				// Indica que se han leido datos
				readed = true;
			}
			// Incrementa la fila
			_row++;
			// Devuelve el valor que indica si se han leído datos
			return readed;
	}

	/// <summary>
	///		Cierra el archivo
	/// </summary>
	public void Close()
	{
		if (_manager != null)
		{
			// Cierra el archivo
			_manager.Close();
			// y libera los datos
			_manager = null;
		}
	}

	/// <summary>
	///		Obtiene el nombre del campo
	/// </summary>
	public string GetName(int i) => Columns[i].name;

	/// <summary>
	///		Obtiene el nombre del tipo de datos
	/// </summary>
	public string GetDataTypeName(int i) => _recordsValues[i].GetType().Name;

	/// <summary>
	///		Obtiene el tipo de un campo
	/// </summary>
	public Type GetFieldType(int i) => GetValue(i)?.GetType();

	/// <summary>
	///		Obtiene el valor de un campo
	/// </summary>
	public object GetValue(int i) => _recordsValues[i];

	public DataTable GetSchemaTable()
	{
		throw new NotImplementedException();
	}

	public int GetValues(object[] values)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	///		Obtiene un valor bool de un campo
	/// </summary>
	public bool GetBoolean(int i) => GetDataValue<bool>(i);

	/// <summary>
	///		Obtiene un valor byte de un campo
	/// </summary>
	public byte GetByte(int i) => GetDataValue<byte>(i);

	public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	///		Obtiene un valor char de un campo
	/// </summary>
	public char GetChar(int i) => GetDataValue<char>(i);

	public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	///		Obtiene un Guid
	/// </summary>
	public Guid GetGuid(int i) => GetDataValue<Guid>(i);

	/// <summary>
	///		Obtiene un entero de 16
	/// </summary>
	public short GetInt16(int i) => GetDataValue<short>(i);

	/// <summary>
	///		Obtiene un entero de 32
	/// </summary>
	public int GetInt32(int i) => GetDataValue<int>(i);

	/// <summary>
	///		Obtiene un entero largo
	/// </summary>
	public long GetInt64(int i) => GetDataValue<long>(i);

	/// <summary>
	///		Obtiene un valor flotante
	/// </summary>
	public float GetFloat(int i) => GetDataValue<float>(i);

	/// <summary>
	///		Obtiene un valor doble
	/// </summary>
	public double GetDouble(int i) => GetDataValue<double>(i);

	/// <summary>
	///		Obtiene una cadena
	/// </summary>
	public string GetString(int i)
	{
		object value = GetValue(i);

			if (value is string resultValue)
				return resultValue;
			else
				return value?.ToString();
	}

	/// <summary>
	///		Obtiene un valor decimal
	/// </summary>
	public decimal GetDecimal(int i) => GetDataValue<decimal>(i);

	/// <summary>
	///		Obtiene una fecha
	/// </summary>
	public DateTime GetDateTime(int i) => GetDataValue<DateTime>(i);

	public IDataReader GetData(int i)
	{
		throw new NotImplementedException();
	}

	/// <summary>
	///		Obtiene un campo de un tipo determinado
	/// </summary>
	private TypeData? GetDataValue<TypeData>(int i)
	{
		object value = GetValue(i);

			if (value is TypeData resultValue)
				return resultValue;
			else
				return default;
	}

	/// <summary>
	///		Obtiene el índice de un campo a partir de su nombre
	/// </summary>
	public int GetOrdinal(string name)
	{
		// Obtiene el índice del registro
		if (!string.IsNullOrWhiteSpace(name))
			for (int index = 0; index < Columns.Count; index++)
				if (Columns[index].name.Equals(name, StringComparison.CurrentCultureIgnoreCase))
					return index;
		// Si ha llegado hasta aquí es porque no ha encontrado el campo
		return -1;
	}

	/// <summary>
	///		Indica si el campo es un DbNull
	/// </summary>
	public bool IsDBNull(int index) => index >= _recordsValues.Count || _recordsValues[index] == null || _recordsValues[index] is DBNull;

	/// <summary>
	///		Los CSV sólo devuelven un Resultset, de todas formas, DbDataAdapter espera este valor
	/// </summary>
	public bool NextResult() => false;

	/// <summary>
	///		Libera la memoria
	/// </summary>
	protected virtual void Dispose(bool disposing)
	{
		if (!IsDisposed)
		{
			// Libera los datos
			if (disposing)
				Close();
			// Indica que se ha liberado
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
	///		Profundidad del recordset
	/// </summary>
	public int Depth => 0;

	/// <summary>
	///		Indica si está cerrado
	/// </summary>
	public bool IsClosed => _manager is null;

	/// <summary>
	///		Registros afectados
	/// </summary>
	public int RecordsAffected => -1;

	/// <summary>
	///		Bloque de filas para las que se lanza el evento de grabación
	/// </summary>
	public int NotifyAfter { get; }

	/// <summary>
	///		Número de campos a partir de las columnas
	/// </summary>
	/// <remarks>
	///		Lo primero que hace un BulkCopy es ver el número de campos que tiene, si no se ha leido la cabecera puede
	///	que aún no tengamos ningún número de columnas, por eso se lee por primera vez
	/// </remarks>
	public int FieldCount => Columns.Count; 

	/// <summary>
	///		Indizador por número de campo
	/// </summary>
	public object this[int i] => _recordsValues[i];

	/// <summary>
	///		Indizador por nombre de campo
	/// </summary>
	public object this[string name]
	{ 
		get 
		{ 
			int index = GetOrdinal(name);

				if (index >= _recordsValues.Count)
					return null;
				else
					return _recordsValues[GetOrdinal(name)]; 
		}
	}

	/// <summary>
	///		Columnas
	/// </summary>
	public List<(string name, Type type)> Columns { get; } = [];

	/// <summary>
	///		Indica si tiene cabeceras
	/// </summary>
	public bool WithHeader { get; }

	/// <summary>
	///		Indica si se ha liberado el recurso
	/// </summary>
	public bool IsDisposed { get; private set; }
}
