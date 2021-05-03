using System;
using ClosedXML.Excel; //Importando a biblioteca ClosedXML.

namespace Gerenciadorxml
{
    class Program
    {
        static void Main(string[] args)
        {
            //Abrindo o arquivo do excel.
            var Ex = new XLWorkbook(@"C:/Produtividade 2019-2020.xlsx");

            //Criando um for para percorrer todas as abas da planilha. 
            for (int T = 1; T >= 1; T++)
            {
                // Planilha recebe todas abas que tiver existente na planilha
                var planilha = Ex.Worksheet(T);

                //Nome da aba da planilha.
                Console.WriteLine("Aba: " + planilha);
                Console.WriteLine();

                //Acessando dados da linha 1 da planilha.
                var linha = 1;

                //Enquanto for verdadeiro, faca...
                while (true)
                {   //nome recebe dados da coluna A da linha 1.
                    var nome = planilha.Cell("A" + linha.ToString()).Value.ToString();

                    //Se encontrar linha vazia, pode encerrar o excel.
                    if (string.IsNullOrEmpty(nome)) break;

                    //Mostrando os dados.
                    Console.Write(nome.PadRight(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("B" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("C" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("D" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("E" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("F" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("G" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("H" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("I" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("J" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("L" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("M" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("N" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("O" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("P" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("Q" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("R" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("S" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("T" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("U" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("V" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("X" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("Y" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Mostrando dados da coluna B.
                    Console.Write(" " + planilha.Cell("Z" + linha.ToString()).Value.ToString().PadLeft(5));

                    //Repetir para a linha 2,3,4 e etc...
                    linha++;
                }
                planilha.Clear();
            }
            Ex.Dispose();
            Console.ReadKey();
        }
    }
}
