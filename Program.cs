using System;
using ClosedXML.Excel; //Importando a biblioteca ClosedXML.

namespace Gerenciadorxml
{
    class Program
    {
        static void Main(string[] args)
        {
            //Abrindo o arquivo do excel.
            var Ex = new XLWorkbook(@"C:/PlanilhaTeste.xlsx");
            //Acessando a aba da planilha.  
            var planilha = Ex.Worksheet(1);
            //Nome da aba da planilha.
            Console.WriteLine("Aba: "+planilha);

            //Acessando dados da linha 1 da planilha.
            var linha = 1;
            //Enquanto for verdadeiro, faca...
            while (true)
            {   //nome recebe dados da coluna A da linha 1.
                var nome = planilha.Cell("A"+ linha.ToString()).Value.ToString();
                //Se encontrar campos vazio na planilha, então pode encerrar.
                if (string.IsNullOrEmpty(nome)) break;
                //Mostrando os dados da linha A.
                Console.Write(nome);
                //Mostrando dados da coluna B da linha 1.
                Console.WriteLine(planilha.Cell("B" + linha.ToString()).Value.ToString());
                //Repetir para a linha 2,3,4 e etc...
                linha++;
            }

            
        }
    }
}
