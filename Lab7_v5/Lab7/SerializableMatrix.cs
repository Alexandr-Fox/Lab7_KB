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
		string[,] array = new string[100,100];

		public SerializableMatrix()
		{}
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
