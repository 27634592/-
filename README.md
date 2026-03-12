# -dotnet new wpf -n MailClient
MailClient/
 ├── App.xaml
 ├── App.xaml.cs
 ├── MainWindow.xaml
 ├── MainWindow.xaml.cs
 ├── MailClient.csproj
 ├── Models/
 │     └── Account.cs
 ├── Services/
 │     ├── AuthService.cs
 │     ├── MailService.cs
 │     ├── FileService.cs
 │     └── ExportService.cs
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net6.0-windows</TargetFramework>
    <UseWPF>true</UseWPF>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Graph" Version="5.41.0" />
    <PackageReference Include="Microsoft.Identity.Client" Version="4.61.0" />
  </ItemGroup>

</Project>
<Application x:Class="MailClient.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             StartupUri="MainWindow.xaml">
    <Application.Resources>
    </Application.Resources>
</Application>
namespace MailClient
{
    public partial class App : Application
    {
    }
}
<Window x:Class="MailClient.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="微软邮箱取件工具" Height="500" Width="800">

    <Grid Margin="10">
        <StackPanel Orientation="Horizontal" Margin="0 0 0 10">
            <Button Content="导入 TXT" Width="100" Click="ImportTxt_Click"/>
            <Button Content="添加邮箱" Width="100" Margin="10 0 0 0" Click="AddEmail_Click"/>
            <Button Content="删除选中" Width="100" Margin="10 0 0 0" Click="DeleteSelected_Click"/>
            <Button Content="取件" Width="100" Margin="10 0 0 0" Click="FetchMail_Click"/>
            <Button Content="导出邮件到 TXT" Width="150" Margin="10 0 0 0" Click="ExportTxt_Click"/>
        </StackPanel>

        <ListBox x:Name="EmailList" SelectionMode="Extended" DisplayMemberPath="Email"/>
    </Grid>
</Window>
using MailClient.Models;
using MailClient.Services;
using Microsoft.Graph.Models;
using System.Windows;

namespace MailClient
{
    public partial class MainWindow : Window
    {
        private List<Account> accounts = new();
        private AuthService auth = new("你的ClientID");
        private MailService mail = new();
        private FileService file = new();
        private ExportService exporter = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void RefreshList()
        {
            EmailList.ItemsSource = null;
            EmailList.ItemsSource = accounts;
        }

        private void ImportTxt_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "文本文件|*.txt";

            if (dlg.ShowDialog() == true)
            {
                var list = file.LoadFromTxt(dlg.FileName);
                accounts.AddRange(list);
                RefreshList();
            }
        }

        private void AddEmail_Click(object sender, RoutedEventArgs e)
        {
            var email = Microsoft.VisualBasic.Interaction.InputBox("输入邮箱：");
            if (!string.IsNullOrEmpty(email))
            {
                accounts.Add(new Account { Email = email });
                RefreshList();
            }
        }

        private void DeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = EmailList.SelectedItems.Cast<Account>().ToList();
            foreach (var acc in selected)
                accounts.Remove(acc);

            RefreshList();
        }

        private async void FetchMail_Click(object sender, RoutedEventArgs e)
        {
            foreach (var acc in EmailList.SelectedItems.Cast<Account>())
            {
                acc.AccessToken = await auth.LoginAsync();

                var inbox = await mail.GetInboxAsync(acc.AccessToken);
                var junk = await mail.GetJunkAsync(acc.AccessToken);

                MessageBox.Show($"{acc.Email}\n收件箱：{inbox.Count} 封\n垃圾箱：{junk.Count} 封");
            }
        }

        private async void ExportTxt_Click(object sender, RoutedEventArgs e)
        {
            var selected = EmailList.SelectedItems.Cast<Account>().ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("请先选择要导出的邮箱");
                return;
            }

            foreach (var acc in selected)
            {
                if (string.IsNullOrEmpty(acc.AccessToken))
                    acc.AccessToken = await auth.LoginAsync();

                var inbox = await mail.GetInboxAsync(acc.AccessToken);
                var junk = await mail.GetJunkAsync(acc.AccessToken);

                var dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.Filter = "文本文件|*.txt";
                dlg.FileName = $"{acc.Email}_邮件导出.txt";

                if (dlg.ShowDialog() == true)
                {
                    exporter.ExportMessagesToTxt(dlg.FileName, acc.Email, inbox, junk);
                    MessageBox.Show($"{acc.Email} 导出成功！");
                }
            }
        }
    }
}
namespace MailClient.Models
{
    public class Account
    {
        public string Email { get; set; }
        public string AccessToken { get; set; }
    }
}
using Microsoft.Identity.Client;

namespace MailClient.Services
{
    public class AuthService
    {
        private readonly IPublicClientApplication _app;

        public AuthService(string clientId)
        {
            _app = PublicClientApplicationBuilder
                .Create(clientId)
                .WithRedirectUri("http://localhost")
                .Build();
        }

        public async Task<string> LoginAsync()
        {
            string[] scopes = { "Mail.Read" };

            var result = await _app.AcquireTokenInteractive(scopes).ExecuteAsync();
            return result.AccessToken;
        }
    }
}
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;

namespace MailClient.Services
{
    public class MailService
    {
        private GraphServiceClient GetClient(string token)
        {
            var authProvider = new BaseBearerTokenAuthenticationProvider(token);
            return new GraphServiceClient(authProvider);
        }

        public async Task<List<Message>> GetInboxAsync(string token)
        {
            var client = GetClient(token);
            var messages = await client.Me.MailFolders["Inbox"].Messages.GetAsync();
            return messages.Value.ToList();
        }

        public async Task<List<Message>> GetJunkAsync(string token)
        {
            var client = GetClient(token);
            var messages = await client.Me.MailFolders["JunkEmail"].Messages.GetAsync();
            return messages.Value.ToList();
        }
    }
}
using MailClient.Models;

namespace MailClient.Services
{
    public class FileService
    {
        public List<Account> LoadFromTxt(string path)
        {
            var list = new List<Account>();

            foreach (var line in File.ReadAllLines(path))
            {
                var parts = line.Split("----");
                if (parts.Length >= 1)
                {
                    list.Add(new Account
                    {
                        Email = parts[0]
                    });
                }
            }

            return list;
        }
    }
}
using Microsoft.Graph.Models;

namespace MailClient.Services
{
    public class ExportService
    {
        public void ExportMessagesToTxt(string filePath, string email, List<Message> inbox, List<Message> junk)
        {
            using var writer = new StreamWriter(filePath, false);

            writer.WriteLine($"邮箱：{email}");
            writer.WriteLine($"导出时间：{DateTime.Now}");
            writer.WriteLine("=====================================");
            writer.WriteLine("============== 收件箱 ===============");
            writer.WriteLine("=====================================");

            foreach (var msg in inbox)
            {
                writer.WriteLine($"发件人：{msg.From?.EmailAddress?.Address}");
                writer.WriteLine($"主题：{msg.Subject}");
                writer.WriteLine($"时间：{msg.ReceivedDateTime}");
                writer.WriteLine($"正文预览：{msg.BodyPreview}");
                writer.WriteLine("-------------------------------------");
            }

            writer.WriteLine();
            writer.WriteLine("=====================================");
            writer.WriteLine("============== 垃圾箱 ===============");
            writer.WriteLine("=====================================");

            foreach (var msg in junk)
            {
                writer.WriteLine($"发件人：{msg.From?.EmailAddress?.Address}");
                writer.WriteLine($"主题：{msg.Subject}");
                writer.WriteLine($"时间：{msg.ReceivedDateTime}");
                writer.WriteLine($"正文预览：{msg.BodyPreview}");
                writer.WriteLine("-------------------------------------");
            }
        }
    }
}
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
