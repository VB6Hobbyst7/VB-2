Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OracleClient
Imports vb = Microsoft.VisualBasic



Public Class frmMain

#Region " ini파일 "
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal IpApplicationName As String, ByVal ipkeyname As String, ByVal ipdefault As String, _
         ByVal ipreturnedstring As String, ByVal nsize As Integer, ByVal ipfilename As String) As Integer
#End Region

#Region " ini파일 정보 읽기 "
    Public Function INI_READ(ByVal session As String, ByVal keyValue As String, ByVal inifile As String) As String
        Dim strBuffer As String
        strBuffer = Space(255)
        GetPrivateProfileString(session, keyValue, "", strBuffer, 255, inifile)
        Return strBuffer
    End Function
#End Region

#Region " 전역 변수 선언 "
    Private strDBConnStr As String = "Data Source={0};User ID={1};Password={2}"
    Private strDB_Base As String
    Private strDB_User As String
    Private strDB_Pass As String

    Private oraCommand As OracleCommand
    Private oraAdapter As OracleDataAdapter
    Private oraConnection As OracleConnection   '오라클커넥션
    Private oraTran As OracleTransaction
#End Region

#Region " 폼 로드 "
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Interval = 50000
        Timer1.Enabled = True

        Call subTotalDisplasy()
    End Sub
#End Region

#Region " 화면 표시"
    Public Sub subRDCPDisplay(ByVal oraconnection As OracleConnection, ByVal dttGridView As DataGridView, ByVal strFrom As String)
        Dim strSQL As New System.Text.StringBuilder
        Dim dttData As DataTable
        Dim intDelayCount As Integer = 0
        Dim LongTimeDelay As Long = 0
        Dim strMsg As String = ""

        Try
            strSQL.Append(" SELECT b.station_name as 관측소명,MAX(a.OBS_TIME) as 관측시간")
            strSQL.Append(" FROM " & strFrom & " a, TK_USN_STATION b")
            strSQL.Append(" WHERE a.SYSTEM_ID = b.SYSTEM_id")
            strSQL.Append(" GROUP BY b.station_name, a.SYSTEM_ID, a.NODE_ID")
            strSQL.Append(" ORDER BY a.SYSTEM_ID")

            dttData = funcSelectDataTable(strSQL.ToString, oraconnection)

            dttGridView.DataSource = dttData

            Call subDataViewSetting(dttGridView)

        Catch ex As Exception
            subLogWrite("subRDCPDisplay : " + ex.ToString)
        End Try

    End Sub

#End Region

#Region " 데이터뷰그리드뷰 셋팅 "
    Public Sub subDataViewSetting(ByVal dttView As DataGridView)

        Dim row As DataGridViewRow

        With dttView
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToOrderColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToResizeColumns = False
            .MultiSelect = False
            .RowHeadersVisible = False
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .Columns(0).Width = 200
            .Columns(1).Width = 150
            .RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken


            For Each c As DataGridViewColumn In dttView.Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next c


            For icnt As Integer = 0 To .Rows.Count - 1

                row = .Rows(icnt)
                row.Height = 16

                If (-1 < DateDiff(DateInterval.Minute, dttView.Item(1, icnt).Value, Now)) AndAlso _
                     (DateDiff(DateInterval.Minute, dttView.Item(1, icnt).Value, Now) < 30) Then
                    .Rows(icnt).DefaultCellStyle.BackColor = Color.White
                    .Rows(icnt).DefaultCellStyle.ForeColor = Color.Black

                    '(수질분석쪽은 1시간 타임이기에 5분을 더 줘서 65분으로 수정 2012/05/08 안재구)
                ElseIf (29 < DateDiff(DateInterval.Minute, dttView.Item(1, icnt).Value, Now)) AndAlso _
                    (DateDiff(DateInterval.Minute, dttView.Item(1, icnt).Value, Now) <= 65) Then
                    .Rows(icnt).DefaultCellStyle.BackColor = Color.Yellow
                    .Rows(icnt).DefaultCellStyle.ForeColor = Color.Black
                Else
                    .Rows(icnt).DefaultCellStyle.BackColor = Color.Red
                    .Rows(icnt).DefaultCellStyle.ForeColor = Color.White

                End If
            Next

            .Height = ((.Rows.Count + 1) * 16) + 3

            .ClearSelection()

        End With
    End Sub
#End Region

#Region " ini파일 정보 취득 "

    Public Sub subGetiniFileInfo()
        Dim strPathINI As String = ""
        Dim strTemp As String = ""

        'ini파일경로
        strPathINI = funcGetAppPath()
        strPathINI = strPathINI + "\config.ini"
        'DB 설정 정보
        strDB_Base = INI_READ("DATABASE", "dataSource", strPathINI)
        strDB_Base = Trim(strDB_Base)
        strDB_Base = strDB_Base.Substring(0, Len(strDB_Base) - 1)

        strDB_User = INI_READ("DATABASE", "userid", strPathINI)
        strDB_User = Trim(strDB_User)
        strDB_User = strDB_User.Substring(0, Len(strDB_User) - 1)

        strDB_Pass = INI_READ("DATABASE", "password", strPathINI)
        strDB_Pass = Trim(strDB_Pass)
        strDB_Pass = strDB_Pass.Substring(0, Len(strDB_Pass) - 1)

    End Sub
#End Region

#Region " 실행파일 PATH취득 "

    Public Function funcGetAppPath() As String
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

#End Region

#Region " 로그파일 작성 "
    Public Sub subLogWrite(ByVal strLogMsg As String, Optional ByVal strEXEName As String = "")

        Dim strDirPath As String        '로그파일경로
        Dim strFilePath As String       '로그파일명
        Dim strLogWrite As StreamWriter

        '파일경로설정
        If strEXEName = "" Then
            strDirPath = funcGetAppPath() + "\Logs\"
        Else
            strDirPath = funcGetAppPath() + "\Logs\" + strEXEName
        End If

        '파일명설정
        strFilePath = strDirPath + "\" + DateTime.Now.ToString("yyyyMMdd") + ".txt"

        '로그폴더생성
        If Not Directory.Exists(strDirPath) Then
            Directory.CreateDirectory(strDirPath)
        End If

        '로그파일열기
        strLogWrite = New StreamWriter(strFilePath, True, System.Text.Encoding.Default)

        '로그파일작성
        strLogWrite.WriteLine(Format(Now, "yyyy-MM-dd HH:mm:ss") + ">>" + strLogMsg)
        strLogWrite.Close()

    End Sub
#End Region

#Region "화면표시"
    Public Sub subTotalDisplasy()
        Try

            'INI파일 정보 취득
            Call subGetiniFileInfo()

            '오라클 연결
            oraConnection = New OracleConnection
            oraConnection.ConnectionString = String.Format(strDBConnStr, strDB_Base.ToString, strDB_User.ToString, strDB_Pass.ToString)
            oraConnection.Open()

            Me.Cursor = Cursors.WaitCursor

            'RDCP부이자료 표시
            Call subRDCPDisplay(oraConnection, dgv_RDCP, "TK_BUOY_OBSERVATION")

            '기상자료 표시
            Call subRDCPDisplay(oraConnection, dgv_weather, "TK_WEATHER_OBSERVATION")

            '조위자료 표시
            Call subRDCPDisplay(oraConnection, dgv_Tidal, "TK_TIDAL_OBSERVATION")

            '수질모니터링자료 표시
            Call subRDCPDisplay(oraConnection, dgv_WQ, "TK_WQ_OBSERVATION")

            '수질분석자료 표시
            Call subRDCPDisplay(oraConnection, dgv_WQA, "TK_WQA_OBSERVATION")

            'CT String자료 표시
            Call subRDCPDisplay(oraConnection, dgv_CT, "TK_BUOY_CTD")

            oraConnection.Close()
            oraConnection = Nothing

            Me.Cursor = Cursors.Default

        Catch ex As Exception
            Call subLogWrite("subTotalDisplasy : " + ex.ToString)
        End Try
    End Sub
#End Region

#Region " 화면클리어 "
    Public Sub subDisplayClear()

        dgv_RDCP.DataSource = Nothing

        dgv_weather.DataSource = Nothing


        dgv_Tidal.DataSource = Nothing


        dgv_WQ.DataSource = Nothing


        dgv_WQA.DataSource = Nothing


        dgv_CT.DataSource = Nothing


    End Sub

#End Region

#Region " SELECT - DataTable"
    Public Function funcSelectDataTable(ByVal strSQL As String, ByVal conn As OracleConnection) As DataTable

        Dim dttData As New DataTable

        Try
            '오라클 코맨드 연결
            oraCommand = New OracleCommand(strSQL, conn)
            '오라클 아답터 연결
            oraAdapter = New OracleDataAdapter(oraCommand)
            '데이터테이블 저장
            oraAdapter.Fill(dttData)
        Catch ex As Exception
            subLogWrite("SQL에러 : " + strSQL)
        Finally
            oraCommand.Dispose()
            oraCommand = Nothing

            oraAdapter.Dispose()
            oraAdapter = Nothing
        End Try

        Return dttData
    End Function
#End Region

#Region " INSERT "
    Public Function funcInsertData(ByVal strSQL As String, ByVal conn As OracleConnection, ByVal Tran As OracleTransaction) As Boolean
        Dim dttData As New DataTable
        Dim intinsertCnt As Integer = 0

        Try
            '오라클 코맨드 연결
            oraCommand = New OracleCommand(strSQL, conn, Tran)
            '오라클 아답터 연결
            intinsertCnt = oraCommand.ExecuteNonQuery()
            '데이터테이블 저장
        Catch ex As Exception
            subLogWrite("SQL에러 : " + strSQL)
            subLogWrite(ex.ToString)
            Tran.Rollback()
            Return False
        Finally
            oraCommand.Dispose()
            oraCommand = Nothing
        End Try

        Return True

    End Function
#End Region

#Region " 타이머실행 "
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Call subDisplayClear()
        Call subTotalDisplasy()
    End Sub
#End Region

#Region " 화면갱신 "
    Private Sub btnReLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReLoad.Click
        Call subDisplayClear()
        Call subTotalDisplasy()
    End Sub
#End Region

#Region " 종료 "
    Private Sub btnEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnd.Click
        Me.Close()
    End Sub
#End Region

End Class

