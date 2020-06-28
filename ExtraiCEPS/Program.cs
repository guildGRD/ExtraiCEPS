using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;

namespace BuscaCEPs
{
    class Program
    {
        static int xlRow;
        
        static bool ExisteElemento(IWebDriver doraiva, By by)
        {
            try
            {
                doraiva.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        static void PreparaConsulta(IWebDriver doraiva)
        {
            //acessa a pagina para o inicio da busca
            doraiva.Url = "http://www.buscacep.correios.com.br/sistemas/buscacep/BuscaCepEndereco.cfm";
            
            int cont = 1;
            //forma que encontrei de burlar o bloqueior dos correios por forca bruta
            while(!ExisteElemento(doraiva, By.XPath("//span[@class='f8col']//a")))
            {
                System.Threading.Thread.Sleep(30000);
                doraiva.Url = "http://www.buscacep.correios.com.br/sistemas/buscacep/BuscaCepEndereco.cfm";
                Console.WriteLine("Tentativa : " + cont.ToString() + " As: " + DateTime.Now.ToString());
                cont++;
            }
            // clicka no radioButton para pesquisar palavras semelhantes
            doraiva.FindElement(By.XPath("//span[@class='f8col']//a")).Click();
            doraiva.FindElement(By.XPath("//input[@name='semelhante'][@value='S']")).Click();
        }

        static bool CarregaPlanilha(IWebDriver doraiva, Excel.Worksheet xlws, string range)
        {
            //pega a tabela com os dados da busca
            var tabelaCeps = doraiva.FindElements(By.XPath("//table[@class='tmptabela']//tr[position() > 1]"));

            if (tabelaCeps.Count > 0)
            {
                foreach (var linha in tabelaCeps)
                {
                    //pega os dados da linha
                    var logradouro = linha.FindElements(By.TagName("td"))[0].Text.Trim().Replace("\n","").Replace("\r","");
                    var bairro = linha.FindElements(By.TagName("td"))[1].Text.Trim();
                    var localidade = linha.FindElements(By.TagName("td"))[2].Text.Trim();
                    var cep = linha.FindElements(By.TagName("td"))[3].Text.Trim();
                    Console.WriteLine(logradouro + bairro + localidade + cep + DateTime.Now.ToString());

                    //excreve os dados na planilha
                    xlws.Cells[xlRow, 1] = range;
                    xlws.Cells[xlRow, 2] = logradouro;
                    xlws.Cells[xlRow, 3] = bairro;
                    xlws.Cells[xlRow, 4] = localidade;
                    xlws.Cells[xlRow, 5] = cep;
                    xlws.Cells[xlRow, 6] = DateTime.Now.ToString();
                    xlRow++;
                }

                //tratamento para a paginacao do retorno
                if (tabelaCeps.Count >= 50 && ExisteElemento(doraiva, By.XPath("(//div[@style='float:left']//a[contains(. , '[ Pr')])[1]")))
                {
                    doraiva.FindElement(By.XPath("(//div[@style='float:left']//a[contains(. , '[ Pr')])[1]")).Click();
                    CarregaPlanilha(doraiva, xlws, range);                    
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Inicando o Processo de busca... \nAguarde...");
            
            object missValue = System.Reflection.Missing.Value;
            IWebDriver driver = null;
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Workbook xlEntradaBook = null;
            Excel.Worksheet xlEntradaSheet = null;
            Excel.Range xlEntradaRange = null;

            try
            {                
                string cepFaixaIni = "";
                string cepFaixaFim = "";
                string rangeCep = "";

                xlApp = new Excel.Application();
                //Cria uma planilha temporária na memória do computador para o input de dados
                xlWorkBook = xlApp.Workbooks.Add(missValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //abre a planilha de entrada para pegar seus valores
                xlEntradaBook = xlApp.Workbooks.Open(Directory.GetCurrentDirectory() + "\\Lista_de_CEPs.xlsx");
                xlEntradaSheet = xlEntradaBook.Sheets[1];
                xlEntradaRange = xlEntradaSheet.UsedRange;

                //Instancia o driver e abre o navegador em modo headless para um ganho de desempenho
                var chromeOptions = new ChromeOptions();
                chromeOptions.AddArguments("headless");
                driver = new ChromeDriver(chromeOptions);

                //comeca a escrever os dados para o excel
                xlWorkSheet.Cells[1, 1] = "Faixa ";
                xlWorkSheet.Cells[1, 2] = "Logradouro/Nome";
                xlWorkSheet.Cells[1, 3] = "Bairro/Distrito";
                xlWorkSheet.Cells[1, 4] = "Localidade/UF";
                xlWorkSheet.Cells[1, 5] = "CEP";
                xlWorkSheet.Cells[1, 6] = "Data e Hora";
                xlRow = 2; 

                PreparaConsulta(driver);
                IWebDriver testT = null;
                //laco com as linhas do excel de entrada
                for (int i = 2; i <= xlEntradaRange.Rows.Count; i++)
                {
                    //pega os valores do xlsx para busca referente a linha 
                    rangeCep = xlEntradaRange.Cells[i, 1].Value2.ToString();
                    cepFaixaIni = xlEntradaRange.Cells[i, 2].Value2.ToString().Substring(0, 5);
                    cepFaixaFim = xlEntradaRange.Cells[i, 3].Value2.ToString().Substring(0, 5);

                    driver.FindElement(By.Name("relaxation")).SendKeys(cepFaixaIni + "%" + Keys.Enter);
                    //Nota: este sendkey funciona mas por algum motivo exibe um erro no console de 
                    //leitura de valor indefinido 

                    //pega todas as linhas na tabela de retorno,e escreve os dados na planilha de saida
                    if (CarregaPlanilha(driver, xlWorkSheet, rangeCep))
                    {
                        PreparaConsulta(driver);
                        driver.FindElement(By.Name("relaxation")).SendKeys(cepFaixaFim + "%" + Keys.Enter);
                    }
                    else
                    {
                        driver.Navigate().Back();
                        driver.FindElement(By.Name("relaxation")).Clear();
                        driver.FindElement(By.Name("relaxation")).SendKeys(cepFaixaFim + "%" + Keys.Enter);
                    }
                   
                    if (CarregaPlanilha(driver, xlWorkSheet, rangeCep))
                    {                         
                        PreparaConsulta(driver);
                    }
                    else
                    {
                        driver.Navigate().Back();
                        driver.FindElement(By.Name("relaxation")).Clear();
                    }
                }

                //Grava a planilha temporaria
                xlWorkSheet.Columns.AutoFit();
                xlWorkSheet.Rows.AutoFit();
                xlWorkBook.SaveAs("resultado.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, missValue, missValue, missValue, missValue,
                                       Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
                xlWorkBook.Close(true, missValue, missValue);

                //o arquivo foi salvo na pasta Meus Documentos.
                string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Console.WriteLine("Concluído. Verifique em " + path + "\\resultado.xlsx");
            }
            catch (Exception myError)
            {
                throw new Exception(myError.Message);
            }
            finally
            {
                if (driver != null)
                    driver.Quit();
                    driver.Dispose();

                xlEntradaBook.Close(true, missValue, missValue);
                xlApp.Quit();
            }

            string caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if(File.Exists(caminho + "\\resultado.xlsx")) 
                System.Diagnostics.Process.Start(caminho + "\\resultado.xlsx");

            Console.WriteLine("Fim do processo no geral =-=");
            Console.ReadLine();
        }
    }
}
