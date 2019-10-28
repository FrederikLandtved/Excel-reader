using ExcelDataReader;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;

namespace ExcelReader1
{
	public class Program
	{
		static void Main(string[] args)
		{
			Reader reader = new Reader();
			reader.ReadExcel();
			Console.ReadKey();
		}
	}
}
