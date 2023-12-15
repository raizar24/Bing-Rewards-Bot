Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.Support.UI
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Interactions
Imports System.IO
Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports System.Net.Http
Imports System.Xml
Imports OpenQA.Selenium.Chromium
Imports System.Reflection.PortableExecutable
Imports System.Net.Mail

Module Program
    Dim driver As IWebDriver
    Dim options As New ChromeOptions()
    Dim optionsFox As New FirefoxOptions()

    'to dos:
    '- Include sick and vacation leave
    ' 11/12/2023 UPDATE: changed to firefox due to user agent extension does not work on chrome anymore
    ' disabled earning points due to microsoft has 15 mins cooldown
    Sub Main(args As String())
        Dim TryAgain As Boolean = False
        Dim chromePath As String = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        Dim aggrument As String = "--remote-debugging-port=9222 --disable-extensions"
        Dim processInfo As New ProcessStartInfo(chromePath, aggrument)
        options.DebuggerAddress = "localhost:9222"
        If args(0) = "earn" Then
            Process.Start(processInfo)
            Thread.Sleep(2000)
            EarnPointsProceedure(args(1))
        End If

        'If MayPasokBa() = False Or loadXML("EarnOnly") = "yes" Then
        '    Process.Start(processInfo)
        '    Thread.Sleep(2000)
        '    'EarnPointsProceedure()
        '    shutdown()
        '    Exit Sub
        'End If
        'Do While TryAgain = False
        '    Process.Start(processInfo)
        '    Thread.Sleep(2000)
        '    driver = New ChromeDriver(ChromeDriverService.CreateDefaultService(), options, TimeSpan.FromMinutes(5))
        '    TryAgain = visitSunFish()
        '    'TryAgain = True
        'Loop
        ''External BCP, always TW status
        'visitBCP()
        'shutdown()
        'EarnPointsProceedure()
    End Sub
    Private Function MayPasokBa()
        'Returns false if is in the set of string <holiday> or Saturday, Sunday and Holiday
        'Return true does not met the conditions
        Dim currentDate As Date = Date.Today '#6/18/2023 12:00:00 AM# 
        Dim dayOfWeek As DayOfWeek = currentDate.DayOfWeek
        If dayOfWeek = DayOfWeek.Sunday Or dayOfWeek = DayOfWeek.Saturday Then
            Return False
        End If
        Dim holidayList As List(Of Date) = GetHoliday()
        If holidayList.Contains(currentDate) Then
            Return False
        End If
        Return True
    End Function
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

    Private Sub EarnPointsProceedure(ByVal agrs As String)
        ' Try
        Dim service As FirefoxDriverService = FirefoxDriverService.CreateDefaultService()
        service.Port = 4444
        service.FirefoxBinaryPath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        Dim foxoption = loadXML("FoxProfile")
        optionsFox.AddArguments("-start-debugger-server 4444", "-profile " & foxoption)
        Dim url As String = "https://www.bing.com/search?q=a"
        Dim isSearchReady = False
        Dim searchTerms As List(Of KeyValuePair(Of Integer, String))

        searchTerms = GetSearchTerms()

        'PC mode
        driver = New ChromeDriver(ChromeDriverService.CreateDefaultService(), options, TimeSpan.FromMinutes(5))
        CheckSessionAccount()
        driver.Navigate().GoToUrl(url)
        driver.Manage.Window.Minimize()
        Thread.Sleep(2000)
        driver.Manage.Window.Maximize()
        Dim count As Integer = 1
        If agrs.Equals("pc") Or agrs.Equals("both") Then
            For indexOfSearchTerms As Integer = 0 To 29
                BingRewards(searchTerms(indexOfSearchTerms).Value)
                count += 1
                If count = 5 Then
                    timer()
                    count = 1
                End If
            Next
        End If
        CloseProgram("chrome")

        'mobile mode
        driver = New FirefoxDriver(service, optionsFox, TimeSpan.FromMinutes(5))
        CheckSessionAccount()
        driver.Navigate().GoToUrl(url)
        driver.Manage.Window.Minimize()
        Thread.Sleep(2000)
        driver.Manage.Window.Maximize()
        count = 1
        If agrs.Equals("m") Or agrs.Equals("both") Then
            For indexOfSearchTerms As Integer = 30 To 50
                BingRewards(searchTerms(indexOfSearchTerms).Value)
                count += 1
                If count = 4 Then
                    timer()
                    count = 1
                End If
            Next
        End If
        CloseProgram("firefox")
        'shutdown()
        'Catch ex As Exception
        'shutdown()
        ' End Try
    End Sub


    Private Sub CloseProgram(processname As String)
        Dim ProgramProcesses As Process() = Process.GetProcessesByName(processname)
        For Each programprocess As Process In ProgramProcesses
            programprocess.CloseMainWindow()
        Next
    End Sub

    Private Function GetHoliday() As List(Of Date)
        'Splits and convert the set of string into date and returns holiday list declared in setting.xml <holiday> formatted MM/DD
        'if naka table iyan ito ang values checking
        'Date.Today = #10/18/2023 12:00:00 AM# compares below
        '#1/1/2023 12:00:00 AM#	
        Dim holidayList As New List(Of Date)()
        Dim holidayString As String = loadXML("holiday")
        Dim holidayDates As String() = holidayString.Split(","c)

        For Each dateStr As String In holidayDates
            Dim monthDay As String() = dateStr.Split("/"c)
            If monthDay.Length = 2 Then
                Dim month As Integer
                Dim day As Integer
                If Integer.TryParse(monthDay(0), month) AndAlso Integer.TryParse(monthDay(1), day) Then
                    Dim holidayDate As New Date(Date.Now.Year, month, day)
                    holidayList.Add(holidayDate)
                End If
            End If
        Next

        Return holidayList
    End Function
    Private Function visitSunFish()
        'returns True if is successfully timein
        Try
            Dim wait As New WebDriverWait(driver, TimeSpan.FromSeconds(300))
            driver.Navigate().GoToUrl("https://sf.dataon.com/sf6/index.cfm")
            wait.Until(Function(d) d.FindElement(By.TagName("body")).Displayed)

            Dim accountinfo As IWebElement = driver.FindElement(By.Id("txtAccount"))
            accountinfo.SendKeys(loadXML("accountid"))
            Dim usernameInput As IWebElement = driver.FindElement(By.Id("txtUserName"))
            usernameInput.SendKeys(loadXML("username"))
            Dim passwordInput As IWebElement = driver.FindElement(By.Id("txtPassword"))
            passwordInput.SendKeys(loadXML("password"))
            Dim loginButton As IWebElement = driver.FindElement(By.Id("btn-logon"))
            loginButton.Click()

            'monthly change password password
            Thread.Sleep(2000)
            Dim ChangePassword As IReadOnlyCollection(Of IWebElement) = driver.FindElements(By.Id("txtOPassword")) 'PC mode check session
            If ChangePassword.Count > 0 Then
                Dim txtOPassword As IWebElement = driver.FindElement(By.Id("txtOPassword"))
                txtOPassword.SendKeys(loadXML("password"))
                Dim txtPassword As IWebElement = driver.FindElement(By.Id("txtPassword"))
                txtPassword.SendKeys(loadXML("NextPassword"))
                Dim txtCPassword As IWebElement = driver.FindElement(By.Id("txtCPassword"))
                txtCPassword.SendKeys(loadXML("NextPassword"))
                Dim changePasswordButton As IWebElement = driver.FindElement(By.CssSelector("input[value='Change Password']"))
                changePasswordButton.Click()
                saveXML(loadXML("password"), "OldPassword")
                saveXML(loadXML("NextPassword"), "Password")
                Thread.Sleep(2000)
                driver.Navigate().GoToUrl("https://sf.dataon.com/sf6/index.cfm")
            End If

            wait.Until(Function(d) d.Title.Contains("Human Resource"))
            Dim attendanceEntryForm As IWebElement = driver.FindElement(By.Id("menu_2"))
            Dim actions As Actions = New Actions(driver)
            actions.MoveToElement(attendanceEntryForm).Perform()
            Dim link As IWebElement = driver.FindElement(By.XPath("//a[contains(text(),'My Attendance Entry Form')]"))
            Dim hover As Actions = New Actions(driver)
            hover.MoveToElement(link).Perform()
            link.Click()

            System.Threading.Thread.Sleep(10000)
            'kaya ganito pagkakacode kasi naka 'frame' ang contents ng btn-start-time
            'para bang 'website' in a 'website' so that I need to move the Chrome Driver pointer from 'content' to 'formtable'
            Dim contentID As IWebElement = driver.FindElement(By.Id("content"))
            Dim actions22 As Actions = New Actions(driver)
            actions22.MoveToElement(contentID).Perform()
            Dim frmSFBody As IWebElement = driver.FindElement(By.Id("frmSFBody"))
            driver.SwitchTo().Frame(frmSFBody)
            Dim formtable As IWebElement = driver.FindElement(By.Id("formtable"))
            actions22.MoveToElement(formtable).Perform()
            Dim TimeStartButton As IWebElement = driver.FindElement(By.Id("btn-start-time"))
            TimeStartButton.Click()
            System.Threading.Thread.Sleep(10000)
            driver.SwitchTo().Alert.Accept()
            Return True

        Catch e As Exception
            logs(e.Message)
            System.Threading.Thread.Sleep(5000)

            Dim chromeProcesses() As Process = Process.GetProcessesByName("chrome")
            For Each chromeProcess As Process In chromeProcesses
                chromeProcess.CloseMainWindow()
            Next
            Return False
        End Try
    End Function

    Private Sub visitBCP()
        Dim wait As New WebDriverWait(driver, TimeSpan.FromSeconds(60))
        driver.Navigate().GoToUrl("https://forms.office.com/pages/responsepage.aspx?id=HRKfoeGBWEip2HNuJn_Ux01ESyKTgE9CjW5YAgw_a_VUOUZKRk9XQUJJWkJRU0hRUlBWNk9DSTQ2TS4u")
        wait.Until(Function(d) d.Title.Contains("Business Continuity Plan"))
        Dim radioButton3 As IWebElement = driver.FindElement(By.CssSelector("input[name='r4a81fb3a5a4e4d059213046c5f766814'][value='No']"))
        radioButton3.Click()
        Thread.Sleep(1000)
        Dim textBox As IWebElement = driver.FindElement(By.CssSelector("input[placeholder='Enter your answer']"))
        textBox.SendKeys(loadXML("username"))
        Dim radioButton As IWebElement = driver.FindElement(By.CssSelector("input[name='rebef48f74b684ba09d12e452909a809b'][value='TW']"))
        radioButton.Click()
        Dim radioButton2 As IWebElement = driver.FindElement(By.CssSelector("input[name='r535dfcebea364ed190fbbacb599bba86'][value='TW']"))
        radioButton2.Click()
        Dim button As IWebElement = driver.FindElement(By.CssSelector("button[data-automation-id='submitButton']"))
        button.Submit()
    End Sub
    'logs for to check errors
    Private Sub logs(ByVal logMessage As String)
        Using writer As New StreamWriter("logs.txt", True)
            writer.WriteLine(logMessage)
        End Using
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

    'Private Sub BingRewards(ByVal searchKey As Char)
    '    Dim url As String = "https://www.bing.com/search?q="
    '    For i As Integer = 0 To 32
    '        Thread.Sleep(2000)
    '        driver.Navigate().GoToUrl(url & searchKey)
    '        searchKey = ChrW(AscW(searchKey) + 1)
    '        Thread.Sleep(1000)
    '    Next
    'End Sub

    'change into search random string so that bing will not detech similar search
    'patch for unable to earn search through URL 11/11/2023
    'Private Sub BingRewards(limit As Integer)
    '    Dim url As String = "https://www.bing.com/search?q=a"
    '    Dim random As New Random()
    '    Dim searchKey As String
    '    driver.Manage.Window.Minimize()
    '    Thread.Sleep(2000)
    '    driver.Manage.Window.Maximize()
    '    driver.Navigate().GoToUrl(url)
    '    For i As Integer = 0 To limit
    '        Dim length As Integer = 11
    '        searchKey = RandomString(length)

    '        Thread.Sleep(3000)
    '        Dim search As IWebElement = driver.FindElement(By.Id("sb_form_q"))
    '        search.Clear()
    '        search.SendKeys(searchKey)
    '        Dim EnterSearch As IWebElement = driver.FindElement(By.Id("sb_form_go"))
    '        EnterSearch.Click()
    '        Thread.Sleep(2000)
    '    Next
    'End Sub

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

    Private Sub shutdown()
        System.Threading.Thread.Sleep(10000)
        Dim psi As New ProcessStartInfo("shutdown", "/s /f /t 0")
        psi.CreateNoWindow = True
        psi.UseShellExecute = False
        Process.Start(psi)
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
                Dim url As String = $"https://trends.google.com/trends/api/dailytrends?hl=en-US&ed={ [date]}&geo=PH&ns=15"
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

                Thread.Sleep(New Random().Next(3000, 5000)) 'may bug kaya nag random ang ginawakong sleep
            Catch ex As HttpRequestException
                Console.WriteLine("Error retrieving Google Trends JSON.")
            Catch ex As KeyNotFoundException
                Console.WriteLine("Cannot parse, JSON keys are modified.")
            End Try
        Next

        ' may mga duplicate so nag distinct code ako
        searchTerms = searchTerms.Distinct().ToList()
        Console.WriteLine($"# of search items: {searchTerms.Count}" & vbCrLf)
        Return searchTerms.Select(Function(term, index) New KeyValuePair(Of Integer, String)(index, term)).ToList()
    End Function





End Module
