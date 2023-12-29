Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium
Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports System.Net.Http
Imports System.Xml
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports OpenQA.Selenium.Interactions
Imports HtmlAgilityPack
Imports Desktop.Robot.Windows
Imports WindowsInput

Module Program
    Dim driver As IWebDriver
    Dim optionsFox As New FirefoxOptions()

    Sub Main(args As String())
        Dim complete As Boolean = False
        Do While complete = False
            complete = DownloadFirefox()
        Loop
        Dim service As FirefoxDriverService = FirefoxDriverService.CreateDefaultService()
        service.Port = 4444
        service.FirefoxBinaryPath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        Dim url As String = "https://www.bing.com/search?q=a"

        Dim searchTerms As List(Of KeyValuePair(Of Integer, String))
        Dim currentDate As String = DateTime.Now.ToString("MMddyy")
        Dim fileName As String = $"{currentDate}.csv"
        If Not File.Exists(fileName) Then
            searchTerms = GetSearchTerms()
            SaveToCSV(searchTerms, fileName)
        Else
            searchTerms = ReadFromCSV(fileName)
        End If
        Dim Cached As List(Of String) = ReadWordsFromFile()
        Dim count As Integer = 1

        Dim extensionPath As String = DownloadUserAgentSwitcher()
        driver = New FirefoxDriver(service, optionsFox, TimeSpan.FromSeconds(30))
        driver.Navigate().GoToUrl("about:addons")

        Dim buttonElement As IWebElement = driver.FindElement(By.ClassName("more-options-button"))
        buttonElement.Click()
        Thread.Sleep(1000)
        Dim panelListElement As IWebElement = driver.FindElement(By.CssSelector("panel-list[role='menu'][align='left'][valign='bottom']"))
        Dim actions22 As Actions = New Actions(driver)
        actions22.MoveToElement(panelListElement).Perform()
        Thread.Sleep(1000)
        Dim panelItemElement2 As IWebElement = driver.FindElement(By.CssSelector("panel-item[action='install-from-file'][data-l10n-id='addon-install-from-file'][role='presentation'][data-l10n-attrs='accesskey']"))
        panelItemElement2.Click()
        Thread.Sleep(1000)

        'It Just Works -Todd Howard //////////////
        Dim robot As New Robot()
        Dim inputSimulator As New InputSimulator()
        inputSimulator.Keyboard.TextEntry(extensionPath)
        Thread.Sleep(1000)
        robot.KeyPress(Desktop.Robot.Key.Enter)
        Dim actions As New Actions(driver)
        Thread.Sleep(1000)
        actions.SendKeys(Keys.Enter)
        actions.KeyDown(Keys.Alt).SendKeys("a").KeyUp(Keys.Alt).Perform()
        Thread.Sleep(1000)
        actions.KeyUp(Keys.Alt).SendKeys("a").KeyUp(Keys.Alt).Perform()
        actions.KeyDown(Keys.Alt).SendKeys("o").KeyUp(Keys.Alt).Perform()
        Thread.Sleep(1000)
        actions.KeyUp(Keys.Alt).SendKeys("o").KeyUp(Keys.Alt).Perform()

        'proceed deleting desktop mode 
        Dim element As IWebElement = driver.FindElement(By.XPath("//span[@class='category-name' and @data-l10n-id='addon-category-extension']"))
        element.Click()
        Thread.Sleep(1000)
        Dim element2 As IWebElement = driver.FindElement(By.XPath("//h3[@class='addon-name' and @id='user-agent-switcher_ninetailed_ninja-heading']"))
        element2.Click()
        Thread.Sleep(1000)
        Dim element3 As IWebElement = driver.FindElement(By.XPath("//button[@is='named-deck-button' and @deck='details-deck' and @name='preferences' and @data-l10n-id='preferences-addon-button']"))
        element3.Click()

        Dim tabsToPress = 9 ' Adjust as needed
        For i = 1 To tabsToPress
            robot.KeyPress(Desktop.Robot.Key.Tab)
            Thread.Sleep(200)
        Next
        robot.KeyPress(Desktop.Robot.Key.Enter)
        Thread.Sleep(200)

        'delete desktop modes in add-ins
        For i As Integer = 1 To 6
            For j As Integer = 1 To 5
                robot.KeyPress(Desktop.Robot.Key.Tab)
                Thread.Sleep(200)
            Next
            ' Press Enter
            robot.KeyPress(Desktop.Robot.Key.Enter)
            Thread.Sleep(200)
        Next
        'end

        'proceed to navigate extension / addin
        Dim element5 As IWebElement = driver.FindElement(By.XPath("//button[@is='discover-button' and @viewid='addons://discover/' and @class='category' and @role='tab' and @name='discover' and @aria-selected='false' and @data-l10n-id='addon-category-discover-title' and @tabindex='-1' and @title='Recommendations']"))
        element5.Click()
        Thread.Sleep(200)

        Dim tabsToPress2 = 15
        For i = 1 To tabsToPress2
            robot.KeyPress(Desktop.Robot.Key.Tab)
            Thread.Sleep(200)
        Next

        robot.KeyPress(Desktop.Robot.Key.Right)
        Thread.Sleep(200)
        robot.KeyPress(Desktop.Robot.Key.Enter)
        Thread.Sleep(2000)
        robot.KeyPress(Desktop.Robot.Key.Enter)
        Thread.Sleep(200)

        'inside the extension
        Dim tabsToPress3 = 8
        For i = 1 To tabsToPress3
            robot.KeyPress(Desktop.Robot.Key.Tab)
            Thread.Sleep(200)
        Next
        robot.KeyPress(Desktop.Robot.Key.Down)
        Thread.Sleep(2000)
        '////////////////////

        CheckSessionAccount()
        driver.Navigate().GoToUrl(url)
        driver.Manage.Window.Minimize()
        Thread.Sleep(2000)
        driver.Manage.Window.Maximize()
        For indexOfSearchTerms As Integer = 30 To 49
            Dim searchtext As String = searchTerms(indexOfSearchTerms).Value
            If Not Cached.Contains(searchtext) Then
                BingRewards(searchtext)
                WriteWordsToFile(searchtext)
                count += 1
                If count = 6 Then
                    timer()
                    count = 1
                End If
            End If
        Next
        CloseProgram("firefox")
        Thread.Sleep(2000)
        driver.Quit()

        'pc mode
        optionsFox.AddArgument("-safe-mode")
        driver = New FirefoxDriver(service, optionsFox, TimeSpan.FromSeconds(30))
        Thread.Sleep(2000)
        CheckSessionAccount()
        driver.Navigate().GoToUrl(url)
        driver.Manage.Window.Minimize()
        Thread.Sleep(2000)
        driver.Manage.Window.Maximize()
        Dim elements77 As IReadOnlyCollection(Of IWebElement) = driver.FindElements(By.Id("id_a")) 'Mobile mode check session
        If elements77.Count > 0 Then
            Dim element6 As IWebElement = driver.FindElement(By.XPath("//input[@type='submit' and @name='submit' and @id='id_a' and @value='Sign in']"))
            element6.Click()
        End If
        For indexOfSearchTerms As Integer = 0 To 29
            Dim searchtext = searchTerms(indexOfSearchTerms).Value
            If Not Cached.Equals(searchtext) Then
                BingRewards(searchtext)
                WriteWordsToFile(searchtext)
                count += 1
                If count = 5 Then
                    timer()
                    count = 1
                End If
            End If
        Next
    End Sub
    Private Sub timer()
        Dim endTime As DateTime = DateTime.Now.AddMinutes(15).AddSeconds(30)

        Do While DateTime.Now < endTime
            Dim remainingTime As TimeSpan = endTime - DateTime.Now

            If remainingTime.TotalSeconds <= 0 Then
                Console.WriteLine("Time is up!")
                Exit Do
            End If

            Dim minutesLeft As Integer = remainingTime.Minutes
            Dim secondsLeft As Integer = remainingTime.Seconds
            Console.WriteLine("{0:00}:{1:00} remaining", minutesLeft, secondsLeft)
            Threading.Thread.Sleep(1000)
        Loop
    End Sub
    Private Sub CloseProgram(processname As String)
        Dim ProgramProcesses As Process() = Process.GetProcessesByName(processname)
        For Each programprocess As Process In ProgramProcesses
            programprocess.CloseMainWindow()
        Next
    End Sub
    'para madaling palitan ang values like username, password at holidays. para hindi build ng build
    Private Function loadXML(ByVal value As String)
        Dim xDoc As XDocument = XDocument.Load("settings.xml")
        Dim xreturn As String = xDoc.Descendants(value).FirstOrDefault().Value
        Return xreturn
    End Function

    Private Sub saveXML(ByVal newVal As String, ByVal varLoc As String)
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("setting.xml")
        Dim node As XmlNode = xmlDoc.SelectSingleNode("/Variable/" & varLoc)
        node.InnerText = newVal
        xmlDoc.Save("setting.xml")
    End Sub

    Private Function RandomString(length As Integer) As String
        Const chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        Dim random As New Random()
        Dim result As New String(Enumerable.Repeat(chars, length) _
        .Select(Function(s) s(random.Next(s.Length))).ToArray())
        Return result
    End Function
    'change into more human like searches 
    Private Sub BingRewards(searchTerm As String)
        Thread.Sleep(2000)
        Dim search As IWebElement = driver.FindElement(By.Id("sb_form_q"))
        search.Clear()
        Thread.Sleep(1000)
        'search.SendKeys(RandomString(2))
        search.SendKeys(searchTerm.Substring(0, 1))
        Thread.Sleep(1000)
        search.Clear()
        search.SendKeys(RandomString(4))
        Thread.Sleep(1000)
        search.SendKeys(searchTerm.Substring(0, 2))
        Thread.Sleep(1000)
        search.Clear()
        search.SendKeys(searchTerm)
        Thread.Sleep(1000)
        Dim EnterSearch As IWebElement = driver.FindElement(By.Id("sb_form_go"))
        Thread.Sleep(2000)
        EnterSearch.Click()
    End Sub

    Private Sub CheckSessionAccount()
        driver.Navigate().GoToUrl("https://www.bing.com/")
        Thread.Sleep(3000)
        Dim elementsPC As IReadOnlyCollection(Of IWebElement) = driver.FindElements(By.Id("id_n")) 'PC mode check session
        If elementsPC.Count > 0 Then
            Dim style As String = elementsPC(0).GetAttribute("style")
            If style.Contains("none") Then
                Thread.Sleep(3000)
                Dim BingAccountBTN As IWebElement = driver.FindElement(By.Id("id_s"))
                BingAccountBTN.Click()
                LoginProceedure()
            End If
            Exit Sub
        End If

        'mobile procedure
        Thread.Sleep(3000)
        Dim HamburgerMenu As IWebElement = driver.FindElement(By.Id("mHamburger"))
        HamburgerMenu.Click()

        Thread.Sleep(3000)
        Dim HBsignin As IWebElement = driver.FindElement(By.Id("HBSignIn"))
        HBsignin.Click()

        Thread.Sleep(3000)
        Dim elementsMobile As IReadOnlyCollection(Of IWebElement) = driver.FindElements(By.Id("id__4")) 'Mobile mode check session
        If elementsMobile.Count > 0 Then
            Dim SignInbtn As IWebElement = driver.FindElement(By.Id("id__4"))
            SignInbtn.Click()
            LoginProceedure()
        End If

        'to hundle the redirection on Login Procedure
        Dim elementsMobile2 As IReadOnlyCollection(Of IWebElement) = driver.FindElements(By.Id("i0116"))
        If elementsMobile2.Count > 0 Then
            LoginProceedure()
        End If
    End Sub

    Private Sub LoginProceedure()
        Thread.Sleep(3000)
        Dim EmailBing As IWebElement = driver.FindElement(By.Id("i0116"))
        EmailBing.SendKeys(loadXML("email"))
        Thread.Sleep(1000)
        Dim NextButton As IWebElement = driver.FindElement(By.Id("idSIButton9"))
        NextButton.Click()

        Thread.Sleep(3000)
        Dim PasswordBing As IWebElement = driver.FindElement(By.Id("i0118"))
        PasswordBing.SendKeys(loadXML("bpassword"))
        Thread.Sleep(1000)
        Dim SignButton As IWebElement = driver.FindElement(By.Id("idSIButton9"))
        SignButton.Click()

        Thread.Sleep(3000)
        Dim OptionalElement As IReadOnlyCollection(Of IWebElement) = driver.FindElements(By.Id("idSIButton9"))
        If OptionalElement.Count > 0 Then
            Dim YesButton As IWebElement = driver.FindElement(By.Id("idSIButton9"))
            Thread.Sleep(1000)
            YesButton.Click()
        End If
    End Sub

    Function GetDates() As List(Of String)
        Dim dates As New List(Of String)()
        For i As Integer = 0 To 3
            ' get dates
            Dim [date] As Date = Date.Now - TimeSpan.FromDays(i)
            'Dim [date] As Date = #11/01/2023 07:33:24 PM# - TimeSpan.FromDays(i)
            ' append in year month date format
            dates.Add([date].ToString("yyyyMMdd"))
        Next
        Return dates
    End Function

    Function GetSearchTerms() As List(Of KeyValuePair(Of Integer, String))
        Dim dates As List(Of String) = GetDates()
        Dim searchTerms As New List(Of String)()

        For Each [date] As String In dates
            Try
                ' get URL
                Dim url As String = $"https://trends.google.com/trends/api/dailytrends?hl=en-US&ed={ [date]}&geo=" & loadXML("region") & "&ns=15"
                Dim client As New HttpClient()
                Dim response As String = client.GetStringAsync(url).Result
                response = response.Substring(5) 'remove first 5 char
                Dim jsonResponse As JObject = JObject.Parse(response)
                ' get all trending searches with their related queries
                For Each topic As JObject In jsonResponse("default")("trendingSearchesDays")(0)("trendingSearches")
                    searchTerms.Add(topic("title")("query").ToString().ToLower())
                    For Each relatedTopic As JObject In topic("relatedQueries")
                        searchTerms.Add(relatedTopic("query").ToString().ToLower())
                    Next
                Next
                Thread.Sleep(New Random().Next(3000, 5000))
            Catch ex As HttpRequestException
                Console.WriteLine("Error retrieving Google Trends JSON.")
            Catch ex As KeyNotFoundException
                Console.WriteLine("Cannot parse, JSON keys are modified.")
            End Try
        Next
        searchTerms = searchTerms.Distinct().ToList()
        Console.WriteLine($"# of search items: {searchTerms.Count}" & vbCrLf)
        Return searchTerms.Select(Function(term, index) New KeyValuePair(Of Integer, String)(index, term)).ToList()
    End Function

    Function ReadWordsFromFile()
        Dim currentDate As String = DateTime.Now.ToString("MMddyy")
        Dim fileName As String = $"{currentDate}.txt"
        Dim Cached As New List(Of String)
        Try
            Using reader As New StreamReader(fileName)
                Dim lines As String() = reader.ReadToEnd().Split(Environment.NewLine)
                For Each line As String In lines
                    Cached.Add(line)
                Next
            End Using
        Catch ex As FileNotFoundException
            Console.WriteLine($"File '{fileName}' not found.")
        Catch ex As Exception
            Console.WriteLine($"An error occurred: {ex.Message}")
        End Try
        Return Cached.Distinct().ToList()
    End Function

    Sub WriteWordsToFile(words As String)
        Dim currentDate As String = DateTime.Now.ToString("MMddyy")
        Dim fileName As String = $"{currentDate}.txt"
        Try
            If Not File.Exists(fileName) Then
                File.Create(fileName).Dispose()
                Console.WriteLine($"File '{fileName}' created.")
            End If
            Using writer As New StreamWriter(fileName, True)
                writer.WriteLine(words)
            End Using
            Console.WriteLine($"Cache '{words}'.")
        Catch ex As Exception
            Console.WriteLine($"An error occurred: {ex.Message}")
        End Try
    End Sub
    Sub SaveToCSV(searchTerms As List(Of KeyValuePair(Of Integer, String)), fileName As String)
        Using writer As New StreamWriter(fileName)
            For Each kvp As KeyValuePair(Of Integer, String) In searchTerms
                writer.WriteLine($"{kvp.Key},{kvp.Value}")
            Next
        End Using
    End Sub

    Function ReadFromCSV(fileName As String) As List(Of KeyValuePair(Of Integer, String))
        Dim searchTerms As New List(Of KeyValuePair(Of Integer, String))()
        Using parser As New TextFieldParser(fileName)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            ' Skip header
            parser.ReadLine()

            While Not parser.EndOfData
                Dim fields As String() = parser.ReadFields()
                If fields.Length = 2 AndAlso Integer.TryParse(fields(0), New Integer) Then
                    searchTerms.Add(New KeyValuePair(Of Integer, String)(CInt(fields(0)), fields(1)))
                End If
            End While
        End Using
        Return searchTerms
    End Function

    Function DownloadFirefox()
        Try
            Dim downloadUrl As String = "https://download.mozilla.org/?product=firefox-latest&os=win&lang=en-US"
            Dim userprofile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim savePath As String = userprofile & "\Downloads\FirefoxInstaller.exe"
            Using httpClient As New HttpClient()
                Dim responseBytes As Byte() = httpClient.GetByteArrayAsync(downloadUrl).Result
                File.WriteAllBytes(savePath, responseBytes)
            End Using

            Dim installProcess As New Process()
            installProcess.StartInfo.FileName = savePath
            installProcess.StartInfo.Arguments = "/S"
            installProcess.StartInfo.UseShellExecute = False
            installProcess.StartInfo.CreateNoWindow = True
            installProcess.Start()
            installProcess.WaitForExit()
            Return True
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
            Return False
        End Try
    End Function

    Function DownloadUserAgentSwitcher()
        Dim addonPageUrl As String = "https://addons.mozilla.org/en-US/firefox/addon/uaswitcher/"
        Dim xpiUrl As String = GetLatestVersionDownloadLink(addonPageUrl)
        If Not String.IsNullOrEmpty(xpiUrl) Then
            Dim userprofile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim savePath As String = userprofile & "\Downloads\uaswitcher.xpi"

            Using httpClient As New HttpClient()
                Dim responseBytes As Byte() = httpClient.GetByteArrayAsync(xpiUrl).Result
                File.WriteAllBytes(savePath, responseBytes)
            End Using
            Return savePath
        Else
            Return String.Empty
            Console.WriteLine("Failed to fetch the download link for the latest version.")
        End If
    End Function

    Function GetLatestVersionDownloadLink(addonPageUrl As String) As String
        Dim httpClient As New HttpClient()
        Dim html As String = httpClient.GetStringAsync(addonPageUrl).Result

        Dim htmlDocument As New HtmlDocument()
        htmlDocument.LoadHtml(html)

        Dim downloadLinkNode As HtmlNode = htmlDocument.DocumentNode.SelectSingleNode("//a[@class='InstallButtonWrapper-download-link' and contains(@href, 'uaswitcher')]")
        If downloadLinkNode IsNot Nothing Then
            Dim downloadLink As String = downloadLinkNode.GetAttributeValue("href", "")
            Return downloadLink
        Else
            Return String.Empty
        End If
    End Function
End Module
