using System;

namespace Lab7
{	
	[Serializable]
	public class SerializableMatrix
	{
		string[,] _array ;
		public int Column => _array.GetLength(1);
		public int Row => _array.GetLength(0);

		public SerializableMatrix(int c,int r)
		{
			_array  = new string[r,c];
		}
		public string this [int i, int j]
		{
			get
			{
				return _array[i, j];
			}
			set
			{
				_array[i, j] = value;
			}
		}
	}
}
