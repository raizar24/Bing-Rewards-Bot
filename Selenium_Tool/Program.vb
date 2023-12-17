Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium
Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports System.Net.Http
Imports System.Xml
Imports System.IO

Module Program
    Dim driver As IWebDriver
    Dim optionsFox As New FirefoxOptions()

    Sub Main(args As String())
        Dim service As FirefoxDriverService = FirefoxDriverService.CreateDefaultService()
        service.Port = 4444
        service.FirefoxBinaryPath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        Dim foxoption = loadXML("FoxProfile")
        optionsFox.AddArguments("-start-debugger-server 4444", "-profile " & foxoption)
        Dim url As String = "https://www.bing.com/search?q=a"
        Dim searchTerms As List(Of KeyValuePair(Of Integer, String))
        searchTerms = GetSearchTerms()
        Dim Cached As New List(Of String)
        Cached = ReadWordsFromFile()
        Dim count As Integer = 1
        'mobile mode
        driver = New FirefoxDriver(service, optionsFox, TimeSpan.FromSeconds(30))
        Thread.Sleep(2000)
        CheckSessionAccount()
        driver.Navigate().GoToUrl(url)
        driver.Manage.Window.Minimize()
        Thread.Sleep(2000)
        driver.Manage.Window.Maximize()
        count = 1
        If args(0).Equals("m") Or args(0).Equals("both") Then
            For indexOfSearchTerms As Integer = 30 To 49
                Dim searchtext As String = searchTerms(indexOfSearchTerms).Value
                If Not Cached.Contains(searchtext) Then
                    BingRewards(searchtext)
                    WriteWordsToFile(searchtext)
                    count += 1
                    If count = 4 Then
                        timer()
                        count = 1
                    End If
                End If
            Next
        End If
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
        If args(0).Equals("pc") Or args(0).Equals("both") Then
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
        End If
    End Sub
    Private Sub timer()
        Dim startTime As DateTime
        startTime = DateTime.Now.AddMinutes(15).AddSeconds(30)
        Do While startTime > DateTime.Now
            Dim remainingTime As TimeSpan = startTime - DateTime.Now

            If remainingTime.TotalSeconds <= 0 Then
                Console.WriteLine("Time is up!")
                Exit Do
            End If
            Dim minutesLeft As Integer = CInt(remainingTime.TotalMinutes)
            Dim secondsLeft As Integer = CInt(remainingTime.Seconds)
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
        'Const chars As String = "sdfjklsdjfnriuerwqenrmmqwehjxvcnasdfhjklsdfklsjdfklj123456677889ABCDERFer"
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
                Dim url As String = $"https://trends.google.com/trends/api/dailytrends?hl=en-US&ed={ [date]}&geo=US&ns=15"
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

End Module
