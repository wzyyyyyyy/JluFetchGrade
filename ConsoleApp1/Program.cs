using Microsoft.VisualBasic;
using OpenQA.Selenium;
using OpenQA.Selenium.DevTools;
using OpenQA.Selenium.Edge;
using System;
using System.Text.RegularExpressions;
String email_ = "";
String passwd_ = "";
String server_key = ""; //为空则不通过server酱推送

void waitLoad(EdgeDriver driver,String xpath)
{
    while (true)
    {
        try
        {
            var eElements = driver.FindElements(By.XPath(xpath));
            if (eElements.Count > 0)
            {
                Thread.Sleep(5000);
                break;
            }
        }
        catch
        {
              break;
        }
        Console.WriteLine("网页未成功加载，等待500ms...");
        Thread.Sleep(500);
    }
}

var service = EdgeDriverService.CreateDefaultService();
EdgeDriver driver = new EdgeDriver(service);

driver.Url = "https://iedu.jlu.edu.cn/";

var email = driver.FindElement(By.XPath("//input[@id='un']"));
email.Click();
email.SendKeys(email_);

var passwd = driver.FindElement(By.XPath("//input[@id='pd']"));
passwd.Click();
passwd.SendKeys(passwd_);

var loginBtn = driver.FindElement(By.XPath("//*[@id=\"index_login_btn\"]/input"));
loginBtn.Click();

waitLoad(driver, "/html/body/main/article/section[2]/div[1]/div/div/div[2]/div/div/div/div[1]");
var cjcxBtn = driver.FindElement(By.XPath("/html/body/main/article/section[2]/div[1]/div/div/div[2]/div/div/div/div[1]"));
cjcxBtn.Click();

Thread.Sleep(5000);
var handles = driver.WindowHandles;
var handle = driver.CurrentWindowHandle;

foreach (var i in handles)
{
    if (i !=  handle)
        driver.SwitchTo().Window(i);
}

int count = 0;

int CountOccurrences(string input, string target)
{
    int count = 0;
    int index = 0;

    while ((index = input.IndexOf(target, index)) != -1)
    {
        index += target.Length;
        count++;
    }

    return count;
}

void postMessage(string msg)
{

    if (!string.IsNullOrEmpty(server_key))
    {
        HttpClient client = new();
        client.GetAsync($"https://sctapi.ftqq.com/{server_key}.send?title=新成绩&desp=" + msg + "&channel=18|9");
    }
}


int t = 1;
while (true)
{
    try
    {
        driver.Navigate().Refresh();
        waitLoad(driver, "//*[@id=\"tabledqxq-index-table\"]");
        var table = driver.FindElement(By.XPath("//*[@id=\"tabledqxq-index-table\"]"));
        int tmp = CountOccurrences(table.Text, "详情");
        string msg = string.Empty;        Console.WriteLine($"正在进行第{t}次查询");
        t++;
        if (count != tmp)
        {
            
            count = tmp;
            for (int i = 0; i < count; i++)
            {
                var score = driver.FindElement(By.XPath($"//*[@id=\"row{i}dqxq-index-table\"]/td[2]/div/span")).Text;
                var subject = driver.FindElement(By.XPath($"//*[@id=\"row{i}dqxq-index-table\"]/td[4]/span")).GetAttribute("title");
                msg += $"{subject}:{score}\n\n\n\n";
                Console.WriteLine($"{subject}:{score}\n");
            }
            msg += $"共查询到{count}科成绩";
            Console.WriteLine($"共查询到{count}科成绩");
            postMessage(msg);
        }
        Thread.Sleep(TimeSpan.FromSeconds(10));
    }
    catch(Exception e)
    {
        Console.WriteLine(e.ToString());
    }
}