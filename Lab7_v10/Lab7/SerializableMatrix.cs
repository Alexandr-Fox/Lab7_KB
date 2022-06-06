using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab7
{	
	[Serializable]
	public class SerializableMatrix
	{
		string[,] array ;
		public int Column => array.GetLength(1);
		public int Row => array.GetLength(0);

		public SerializableMatrix(int c,int r)
		{
			array  = new string[r,c];
		}
		public string this [int i, int j]
		{
			get
			{
				return array[i, j];
			}
			set
			{
				array[i, j] = value;
			}
		}
	}
}
