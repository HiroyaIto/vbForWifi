'****************************************
'     SockT0001 ユーザーコントロール
'            2017-7-23 v3.00
'****************************************
Imports System.Net
Imports System.Net.Sockets
Imports System.ComponentModel
Imports System.Threading

Public Class SockT0001Ctrl
    '****************************************
    '             モジュール定数
    '****************************************
    Private Const BUFF_SIZE_DEFAULT As Integer = 8192

    '****************************************
    '             モジュール変数
    '****************************************
    Private CommSocket As Socket
    Private RecvBuffer() As Byte
    Private SendBuffer() As Byte

    '****************************************
    '               プロパティ
    '****************************************
    Property LocalPort As Integer
    Property RemotePort As Integer
    Property RemoteIP As String
    Property BindIP As String
    Private _State As Integer
    ReadOnly Property State As Integer
        Get
            Return _State
        End Get
    End Property
    Private _IPv6Mode As Integer
    Property IPv6Mode As Integer
        Get
            Return _IPv6Mode
        End Get
        Set(value As Integer)
            If (State <> 0) Then
                RaiseError("IPv6Mode Set 1", "通信中はIPv6Modeを設定できません。")
                Exit Property
            End If
            _IPv6Mode = value
        End Set
    End Property
    Private _SendBuffSize As Integer
    Property SendBuffSize As Integer
        Get
            Return _SendBuffSize
        End Get
        Set(value As Integer)
            If (value < 0) Then
                RaiseError("SendBuffSize Set 1", "送信バッファサイズの値が不正です。")
                Exit Property
            End If
            If (State <> 0) Then
                RaiseError("SendBuffSize Set 2", "通信中は送信バッファサイズを設定できません。")
                Exit Property
            End If
            _SendBuffSize = value
            ReDim SendBuffer(_SendBuffSize)
        End Set
    End Property
    Private _RecvBuffSize As Integer
    Property RecvBuffSize As Integer
        Get
            Return _RecvBuffSize
        End Get
        Set(value As Integer)
            If (value < 0) Then
                RaiseError("RecvBuffSize Set 1", "受信バッファサイズの値が不正です。")
                Exit Property
            End If
            If (State <> 0) Then
                RaiseError("RecvBuffSize Set 2", "通信中は受信バッファサイズを設定できません。")
                Exit Property
            End If
            _RecvBuffSize = value
            ReDim RecvBuffer(_RecvBuffSize)
        End Set
    End Property

    '****************************************
    '                イベント
    '****************************************
    Event Opened(ByVal sender As Object)
    Event Bound(ByVal sender As Object)
    Event StateChanged(ByVal sender As Object)
    Event Listening(ByVal sender As Object)
    Event ConnectionRequest(ByVal sender As Object, ByVal comm_socket As Socket)
    Event Closed(ByVal sender As Object, ByVal disc_state As Integer)
    Event Connected(ByVal sender As Object)
    Event Connecting(ByVal sender As Object)
    Event DataArrival(ByVal sender As Object, ByVal recv_data() As Byte, ByVal recv_bytes As Integer)
    Event SendProgress(ByVal sender As Object, ByVal send_bytes As Integer)
    Event SendComplete(ByVal sender As Object, ByVal send_bytes As Integer)
    Event [Error](ByVal sender As Object, ByVal err_msg As String)

    '****************************************
    '                 各処理
    '****************************************
    Public Sub New()
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        ChangeState(0)
        IPv6Mode = 0
        SendBuffSize = BUFF_SIZE_DEFAULT
        RecvBuffSize = BUFF_SIZE_DEFAULT
    End Sub

    Private Sub RaiseError(ByVal err_title As String, ByVal err_msg As String)
        'RaiseEvent Error(Me, err_title & ":" & err_msg)
        RaiseEventSafe([ErrorEvent], New Object() {Me, err_title & ":" & err_msg})
    End Sub

    '***** マルチスレッド対応のイベント呼び出し処理 *****
    '***** (スレッドをまたいだ呼び出しが必要かどうかをここでチェックしている) *****
    '***** (非同期ソケット通信は、内部で自動的にマルチスレッドになるため本処理が必要) *****
    'Private Sub RaiseEventSafe(ByVal ev As System.Delegate, ByRef args() As Object)
    Private Sub RaiseEventSafe(ByVal ev As System.Delegate, ByVal args() As Object)
        Dim bFired As Boolean
        Dim syncInvoke As ISynchronizeInvoke

        If (ev IsNot Nothing) Then
            For Each singleCast As System.Delegate In ev.GetInvocationList()
                bFired = False
                Try
                    'syncInvoke = CType(singleCast.Target, ISynchronizeInvoke)
                    syncInvoke = DirectCast(singleCast.Target, ISynchronizeInvoke)
                    If ((syncInvoke IsNot Nothing) AndAlso (syncInvoke.InvokeRequired)) Then
                        bFired = True
                        syncInvoke.BeginInvoke(singleCast, args)
                    Else
                        bFired = True
                        singleCast.DynamicInvoke(args)
                    End If
                Catch ex As Exception
                    If (Not bFired) Then singleCast.DynamicInvoke(args)
                End Try
            Next
        End If
    End Sub

    Private Sub ChangeState(ByVal new_state As Integer)
        Dim old_state As Integer

        old_state = _State
        _State = new_state
        If (old_state <> new_state) Then
            'RaiseEvent StateChanged(Me)
            RaiseEventSafe(StateChangedEvent, New Object() {Me})
        End If
    End Sub

    '***** マルチキャストの確認 *****
    Private Function CheckMulticastAddress(ByVal ip As IPAddress) As Boolean
        Dim mcast_flag As Boolean
        Dim ip_byte As Byte()

        mcast_flag = False
        ip_byte = ip.GetAddressBytes
        '***** H.H 変更 2015-11-14 *****
        '***** IPv6対応 *****
        If (IPv6Mode = 0 OrElse
            (IPv6Mode = 2 AndAlso ip.AddressFamily = AddressFamily.InterNetwork)) Then
            If ((ip_byte(0) And &HE0) = &HE0) Then
                mcast_flag = True
            End If
        Else
            If (ip_byte(0) = &HFF) Then
                mcast_flag = True
            End If
        End If
        CheckMulticastAddress = mcast_flag
    End Function

    '***** Anyアドレスの取得 *****
    Private Function GetAnyAddress() As IPAddress
        '***** H.H 変更 2015-11-14 *****
        '***** IPv6対応 *****
        If (IPv6Mode = 0) Then
            GetAnyAddress = IPAddress.Any
        Else
            GetAnyAddress = IPAddress.IPv6Any
        End If
    End Function

    '****************************************
    '                メソッド
    '****************************************

    '***** UDP処理 ここから *****

    Public Sub Open()
        Dim mcast_flag As Boolean
        Dim remote_address As IPAddress
        Dim local_address As IPAddress
        Dim local_ip_byte As Byte()
        Dim local_scopeid_byte As Byte()

        '***** マルチキャストの確認 *****
        Try
            remote_address = IPAddress.Parse(RemoteIP)
            mcast_flag = CheckMulticastAddress(remote_address)
        Catch ex As Exception
        End Try

        Try
            If (CommSocket IsNot Nothing) Then
                Try
                    CommSocket.Shutdown(SocketShutdown.Both)
                Catch ex As Exception
                End Try
                Try
                    CommSocket.Close()
                Catch ex As Exception
                End Try
            End If
            '***** H.H 変更 2015-11-14 *****
            '***** IPv6対応 *****
            If (IPv6Mode = 0) Then
                CommSocket = New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
            Else
                CommSocket = New Socket(AddressFamily.InterNetworkV6, SocketType.Dgram, ProtocolType.Udp)
                If (IPv6Mode = 2) Then
                    CommSocket.SetSocketOption(SocketOptionLevel.IPv6, SocketOptionName.IPv6Only, 0)
                End If
            End If

            '***** H.H 変更 2013-3-2 *****
            '***** バッファサイズ設定をソケットにも反映 *****
            CommSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendBuffer, SendBuffSize)
            CommSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveBuffer, RecvBuffSize)

            '***** マルチキャスト送信対応 *****
            If (mcast_flag = True) Then
                remote_address = IPAddress.Parse(RemoteIP)
                If (BindIP <> "") Then
                    local_address = IPAddress.Parse(BindIP)
                Else
                    local_address = GetAnyAddress()
                End If
                '***** H.H 変更 2015-11-14 *****
                '***** IPv6対応 *****
                If (IPv6Mode = 0 OrElse
                    (IPv6Mode = 2 AndAlso remote_address.AddressFamily = AddressFamily.InterNetwork)) Then
                    local_ip_byte = local_address.GetAddressBytes
                    'Set the required interface for outgoing multicast packets.
                    CommSocket.SetSocketOption(SocketOptionLevel.IP, SocketOptionName.MulticastInterface, local_ip_byte)
                Else
                    local_scopeid_byte = BitConverter.GetBytes(Convert.ToInt32(local_address.ScopeId))
                    'Set the required interface for outgoing multicast packets.
                    CommSocket.SetSocketOption(SocketOptionLevel.IPv6, SocketOptionName.MulticastInterface, local_scopeid_byte)
                End If
            End If

            ChangeState(2)
            'RaiseEvent Opened(Me)
            RaiseEventSafe(OpenedEvent, New Object() {Me})
        Catch ex As SocketException
            RaiseError("Open 1", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("Open 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    Public Sub Bind()
        Dim ipe As IPEndPoint
        Dim mcast_flag As Boolean
        Dim remote_address As IPAddress
        Dim local_address As IPAddress
        Dim mcast_option As MulticastOption
        Dim mcast_option2 As IPv6MulticastOption

        '***** マルチキャストの確認 *****
        Try
            remote_address = IPAddress.Parse(RemoteIP)
            mcast_flag = CheckMulticastAddress(remote_address)
        Catch ex As Exception
        End Try

        Try
            '***** マルチキャストのとき 同一ポート番号のオープンを許可 *****
            If (mcast_flag = True) Then
                CommSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, True)
            End If

            If (BindIP <> "") Then
                ipe = New IPEndPoint(IPAddress.Parse(BindIP), LocalPort)
            Else
                ipe = New IPEndPoint(GetAnyAddress(), LocalPort)
            End If
            CommSocket.Bind(ipe)

            '***** マルチキャスト受信対応 *****
            If (mcast_flag = True) Then
                remote_address = IPAddress.Parse(RemoteIP)
                If (BindIP <> "") Then
                    local_address = IPAddress.Parse(BindIP)
                Else
                    local_address = GetAnyAddress()
                End If
                '***** H.H 変更 2015-11-14 *****
                '***** IPv6対応 *****
                If (IPv6Mode = 0 OrElse
                    (IPv6Mode = 2 AndAlso remote_address.AddressFamily = AddressFamily.InterNetwork)) Then
                    mcast_option = New MulticastOption(remote_address, local_address)
                    'Add membership to the group.
                    CommSocket.SetSocketOption(SocketOptionLevel.IP, SocketOptionName.AddMembership, mcast_option)
                Else
                    mcast_option2 = New IPv6MulticastOption(remote_address, local_address.ScopeId)
                    'Add membership to the group.
                    CommSocket.SetSocketOption(SocketOptionLevel.IPv6, SocketOptionName.AddMembership, mcast_option2)
                End If
            End If

            ipe = New IPEndPoint(GetAnyAddress(), 0)
            'CommSocket.BeginReceiveFrom(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, ipe, New AsyncCallback(AddressOf ReceiveFromCallback), CommSocket)
            ChangeState(2)
            'RaiseEvent Bound(Me)
            RaiseEventSafe(BoundEvent, New Object() {Me})
            CommSocket.BeginReceiveFrom(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, ipe, New AsyncCallback(AddressOf ReceiveFromCallback), CommSocket)
        Catch ex As SocketException
            RaiseError("Bind 1", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("Bind 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    Private Sub ReceiveFromCallback(ByVal iar As IAsyncResult)
        Dim ipe As IPEndPoint
        Dim recv_bytes As Integer
        Dim recv_data() As Byte

        Try
            'CommSocket = DirectCast(iar.AsyncState, Socket)
            ipe = New IPEndPoint(GetAnyAddress(), 0)
            recv_bytes = CommSocket.EndReceiveFrom(iar, ipe)
            If (recv_bytes = 0) Then
                CloseSub(1)
            Else
                ReDim recv_data(recv_bytes - 1)
                Array.Copy(RecvBuffer, 0, recv_data, 0, recv_bytes)
                'RaiseEvent DataArrival(Me, recv_data, recv_bytes)
                RaiseEventSafe(DataArrivalEvent, New Object() {Me, recv_data, recv_bytes})
                ipe = New IPEndPoint(GetAnyAddress(), 0)
                CommSocket.BeginReceiveFrom(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, ipe, New AsyncCallback(AddressOf ReceiveFromCallback), CommSocket)
            End If
        Catch ex As System.ArgumentException
        Catch ex As System.ObjectDisposedException
        Catch ex As SocketException
            RaiseError("ReceiveFromCallback 1", ex.ToString & "(" & ex.ErrorCode & ")")
            'CloseSub(0)
            CloseSub(1)
        Catch ex As Exception
            RaiseError("ReceiveFromCallback 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    '***** H.H 変更 2013-2-16 *****
    '***** 送信データサイズを指定可能にした(メソッドのオーバーロード) *****
    Public Overloads Sub SendDataTo(ByVal send_data() As Byte)
        Dim send_bytes As Integer

        send_bytes = send_data.Length
        SendDataToSub(send_data, send_bytes)
    End Sub
    Public Overloads Sub SendDataTo(ByVal send_data() As Byte, ByVal send_bytes As Integer)
        SendDataToSub(send_data, send_bytes)
    End Sub
    Private Sub SendDataToSub(ByVal send_data() As Byte, ByVal send_bytes As Integer)
        Dim ipe As IPEndPoint
        Dim total_send_bytes As Integer
        Dim will_send_bytes As Integer
        Dim now_send_bytes As Integer
        Dim send_bytes2 As Integer

        Try
            total_send_bytes = send_bytes
            If (total_send_bytes > send_data.Length) Then
                total_send_bytes = send_data.Length
            End If
            If (total_send_bytes < 0) Then
                total_send_bytes = 0
            End If

            '***** H.H 変更 2013-2-4 *****
            '***** 送信バッファサイズを超えた分を捨てないように修正 *****
            'If (total_send_bytes > SendBuffSize) Then
            '    total_send_bytes = SendBuffSize
            'End If
            If (total_send_bytes > SendBuffSize) Then
                will_send_bytes = SendBuffSize
            Else
                will_send_bytes = total_send_bytes
            End If
            If (total_send_bytes > SendBuffer.Length) Then
                ReDim SendBuffer(total_send_bytes - 1)
            End If

            Array.Copy(send_data, 0, SendBuffer, 0, total_send_bytes)
            now_send_bytes = 0
            ipe = New IPEndPoint(IPAddress.Parse(RemoteIP), RemotePort)

            '***** H.H 変更 2013-2-8 *****
            '***** 非同期送信を同期送信に変更 *****
            '***** H.H 変更 2013-2-8 *****
            '***** 0バイトの送信処理変更 *****
            Do

                '***** H.H 変更 2013-2-4 *****
                '***** 送信バッファサイズを超えた分を捨てないように修正 *****
                'send_bytes2 = CommSocket.SendTo(SendBuffer, now_send_bytes, (total_send_bytes - now_send_bytes), SocketFlags.None, ipe)
                send_bytes2 = CommSocket.SendTo(SendBuffer, now_send_bytes, will_send_bytes, SocketFlags.None, ipe)

                If (SendBuffSize > 0) Then
                    now_send_bytes = now_send_bytes + send_bytes2
                    If (now_send_bytes < total_send_bytes) Then
                        'RaiseEvent SendProgress(Me, send_bytes2)
                        RaiseEventSafe(SendProgressEvent, New Object() {Me, send_bytes2})

                        '***** H.H 変更 2013-2-4 *****
                        '***** 送信バッファサイズを超えた分を捨てないように修正 *****
                        If ((total_send_bytes - now_send_bytes) > SendBuffSize) Then
                            will_send_bytes = SendBuffSize
                        Else
                            will_send_bytes = total_send_bytes - now_send_bytes
                        End If

                        Continue Do
                    End If
                End If
                Exit Do
            Loop
            'RaiseEvent SendComplete(Me, send_bytes2)
            RaiseEventSafe(SendCompleteEvent, New Object() {Me, send_bytes2})

        Catch ex As SocketException
            RaiseError("SendDataToSub 1", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("SendDataToSub 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    '***** TCPサーバー処理 ここから *****

    Public Sub Listen()
        Dim ipe As IPEndPoint

        Try
            If (CommSocket IsNot Nothing) Then
                'Try
                '    CommSocket.Shutdown(SocketShutdown.Both)
                'Catch ex As Exception
                'End Try
                Try
                    CommSocket.Close()
                    '***** 待たないと連続Listenでエラー発生することあり *****
                    Thread.Sleep(100)
                Catch ex As Exception
                End Try
            End If
            '***** H.H 変更 2015-11-14 *****
            '***** IPv6対応 *****
            If (IPv6Mode = 0) Then
                CommSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            Else
                CommSocket = New Socket(AddressFamily.InterNetworkV6, SocketType.Stream, ProtocolType.Tcp)
                If (IPv6Mode = 2) Then
                    CommSocket.SetSocketOption(SocketOptionLevel.IPv6, SocketOptionName.IPv6Only, 0)
                End If
            End If
            If (BindIP <> "") Then
                ipe = New IPEndPoint(IPAddress.Parse(BindIP), LocalPort)
            Else
                ipe = New IPEndPoint(GetAnyAddress(), LocalPort)
            End If
            CommSocket.Bind(ipe)
            CommSocket.Listen(100)
            ChangeState(1)
            'RaiseEvent Listening(Me)
            RaiseEventSafe(ListeningEvent, New Object() {Me})
            CommSocket.BeginAccept(New AsyncCallback(AddressOf AcceptCallback), CommSocket)
        Catch ex As SocketException
            RaiseError("Listen 1", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("Listen 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    Private Sub AcceptCallback(ByVal iar As IAsyncResult)
        Dim comm_socket As Socket
        Dim ipe As IPEndPoint

        Try
            'CommSocket = DirectCast(iar.AsyncState, Socket)
            comm_socket = CommSocket.EndAccept(iar)
            ipe = comm_socket.RemoteEndPoint
            RemoteIP = ipe.Address.ToString
            RemotePort = ipe.Port
            'RaiseEvent ConnectionRequest(Me, comm_socket)
            RaiseEventSafe(ConnectionRequestEvent, New Object() {Me, comm_socket})
        Catch ex As System.ArgumentException
        Catch ex As System.ObjectDisposedException
        Catch ex As SocketException
            RaiseError("AcceptCallback 1", ex.ToString & "(" & ex.ErrorCode & ")")
            '***** 連続Connect要求で発生することあり切断すると復旧せず *****
            If (ex.ErrorCode = SocketError.ConnectionReset) Then
                'CloseSub(0)
            Else
                CloseSub(0)
            End If
        Catch ex As Exception
            RaiseError("AcceptCallback 2", ex.ToString)
            CloseSub(0)
        End Try
        Try
            CommSocket.BeginAccept(New AsyncCallback(AddressOf AcceptCallback), CommSocket)
        Catch ex As System.ArgumentException
        Catch ex As System.ObjectDisposedException
        Catch ex As SocketException
            RaiseError("AcceptCallback 3", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("AcceptCallback 4", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    Public Sub Accept(ByVal comm_socket As Socket)
        Dim ipe As IPEndPoint

        Try
            If (CommSocket IsNot Nothing) Then
                Try
                    CommSocket.Shutdown(SocketShutdown.Both)
                Catch ex As Exception
                End Try
                Try
                    CommSocket.Close()
                Catch ex As Exception
                End Try
            End If
            CommSocket = comm_socket

            '***** H.H 変更 2013-3-3 *****
            '***** Acceptの設定抜け修正 *****
            '***** H.H 変更 2013-3-2 *****
            '***** バッファサイズ設定をソケットにも反映 *****
            CommSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendBuffer, SendBuffSize)
            CommSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveBuffer, RecvBuffSize)

            'Dim send_buff_size As Integer
            'Dim recv_buff_size As Integer
            'send_buff_size = BitConverter.ToInt32(CommSocket.GetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendBuffer, 4), 0)
            'recv_buff_size = BitConverter.ToInt32(CommSocket.GetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveBuffer, 4), 0)
            'MessageBox.Show(send_buff_size & " " & recv_buff_size)

            ipe = CommSocket.RemoteEndPoint
            RemoteIP = ipe.Address.ToString
            RemotePort = ipe.Port
            'CommSocket.BeginReceive(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), CommSocket)
            ChangeState(2)
            'RaiseEvent Connected(Me)
            RaiseEventSafe(ConnectedEvent, New Object() {Me})
            CommSocket.BeginReceive(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), CommSocket)
        Catch ex As SocketException
            RaiseError("Accept 1", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("Accept 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    '***** TCPクライアント処理 ここから *****

    Public Sub Connect()
        Dim ipe As IPEndPoint

        Try
            If (CommSocket IsNot Nothing) Then
                Try
                    CommSocket.Shutdown(SocketShutdown.Both)
                Catch ex As Exception
                End Try
                Try
                    CommSocket.Close()
                Catch ex As Exception
                End Try
            End If
            '***** H.H 変更 2015-11-14 *****
            '***** IPv6対応 *****
            If (IPv6Mode = 0) Then
                CommSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            Else
                CommSocket = New Socket(AddressFamily.InterNetworkV6, SocketType.Stream, ProtocolType.Tcp)
                If (IPv6Mode = 2) Then
                    CommSocket.SetSocketOption(SocketOptionLevel.IPv6, SocketOptionName.IPv6Only, 0)
                End If
            End If

            '***** H.H 変更 2013-3-2 *****
            '***** バッファサイズ設定をソケットにも反映 *****
            CommSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendBuffer, SendBuffSize)
            CommSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveBuffer, RecvBuffSize)

            If (BindIP <> "") Then
                ipe = New IPEndPoint(IPAddress.Parse(BindIP), 0)
            Else
                ipe = New IPEndPoint(GetAnyAddress(), 0)
            End If
            CommSocket.Bind(ipe)
            ipe = New IPEndPoint(IPAddress.Parse(RemoteIP), RemotePort)
            'CommSocket.BeginConnect(ipe, New AsyncCallback(AddressOf ConnectCallback), CommSocket)
            ChangeState(1)
            'RaiseEvent Connecting(Me)
            RaiseEventSafe(ConnectingEvent, New Object() {Me})
            CommSocket.BeginConnect(ipe, New AsyncCallback(AddressOf ConnectCallback), CommSocket)
        Catch ex As SocketException
            RaiseError("Connect 1", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("Connect 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    Private Sub ConnectCallback(ByVal iar As IAsyncResult)
        Dim ipe As IPEndPoint

        Try
            'CommSocket = DirectCast(iar.AsyncState, Socket)
            CommSocket.EndConnect(iar)
            ipe = CommSocket.LocalEndPoint
            LocalPort = ipe.Port
            'CommSocket.BeginReceive(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), CommSocket)
            ChangeState(2)
            'RaiseEvent Connected(Me)
            RaiseEventSafe(ConnectedEvent, New Object() {Me})
            CommSocket.BeginReceive(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), CommSocket)
        Catch ex As System.ArgumentException
        Catch ex As System.ObjectDisposedException
        Catch ex As SocketException
            RaiseError("ConnectCallback 1", ex.ToString & "(" & ex.ErrorCode & ")")
            'CloseSub(0)
            CloseSub(1)
        Catch ex As Exception
            RaiseError("ConnectCallback 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    '***** 共通処理 ここから *****

    Private Sub ReceiveCallback(ByVal iar As IAsyncResult)
        Dim recv_bytes As Integer
        Dim recv_data() As Byte

        Try
            'CommSocket = DirectCast(iar.AsyncState, Socket)
            recv_bytes = CommSocket.EndReceive(iar)
            If (recv_bytes = 0) Then
                MsgBox("Off forced")
                CloseSub(1)
            Else
                ReDim recv_data(recv_bytes - 1)
                Array.Copy(RecvBuffer, 0, recv_data, 0, recv_bytes)
                'RaiseEvent DataArrival(Me, recv_data, recv_bytes)
                RaiseEventSafe(DataArrivalEvent, New Object() {Me, recv_data, recv_bytes})
                CommSocket.BeginReceive(RecvBuffer, 0, RecvBuffSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), CommSocket)
            End If
        Catch ex As System.ArgumentException
            MsgBox("Off forced")
        Catch ex As System.ObjectDisposedException
            MsgBox("Off forced")
        Catch ex As SocketException
            RaiseError("ReceiveCallback 1", ex.ToString & "(" & ex.ErrorCode & ")")
            'CloseSub(0)
            CloseSub(1)
        Catch ex As Exception
            RaiseError("ReceiveCallback 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    '***** H.H 変更 2013-2-16 *****
    '***** 送信データサイズを指定可能にした(メソッドのオーバーロード) *****
    Public Overloads Sub SendData(ByVal send_data() As Byte)
        Dim send_bytes As Integer

        send_bytes = send_data.Length
        SendDataSub(send_data, send_bytes)
    End Sub
    Public Overloads Sub SendData(ByVal send_data() As Byte, ByVal send_bytes As Integer)
        SendDataSub(send_data, send_bytes)
    End Sub
    Private Sub SendDataSub(ByVal send_data() As Byte, ByVal send_bytes As Integer)
        Dim total_send_bytes As Integer
        Dim will_send_bytes As Integer
        Dim now_send_bytes As Integer
        Dim send_bytes2 As Integer

        Try
            total_send_bytes = send_bytes
            If (total_send_bytes > send_data.Length) Then
                total_send_bytes = send_data.Length
            End If
            If (total_send_bytes < 0) Then
                total_send_bytes = 0
            End If

            '***** H.H 変更 2013-2-4 *****
            '***** 送信バッファサイズを超えた分を捨てないように修正 *****
            'If (total_send_bytes > SendBuffSize) Then
            '    total_send_bytes = SendBuffSize
            'End If
            If (total_send_bytes > SendBuffSize) Then
                will_send_bytes = SendBuffSize
            Else
                will_send_bytes = total_send_bytes
            End If
            If (total_send_bytes > SendBuffer.Length) Then
                ReDim SendBuffer(total_send_bytes - 1)
            End If

            Array.Copy(send_data, 0, SendBuffer, 0, total_send_bytes)
            now_send_bytes = 0

            '***** H.H 変更 2013-2-8 *****
            '***** 非同期送信を同期送信に変更 *****
            '***** H.H 変更 2013-2-8 *****
            '***** 0バイトの送信処理変更 *****
            Do

                '***** H.H 変更 2013-2-4 *****
                '***** 送信バッファサイズを超えた分を捨てないように修正 *****
                'send_bytes2 = CommSocket.Send(SendBuffer, now_send_bytes, (total_send_bytes - now_send_bytes), SocketFlags.None)
                send_bytes2 = CommSocket.Send(SendBuffer, now_send_bytes, will_send_bytes, SocketFlags.None)

                If (SendBuffSize > 0) Then
                    now_send_bytes = now_send_bytes + send_bytes2
                    If (now_send_bytes < total_send_bytes) Then
                        'RaiseEvent SendProgress(Me, send_bytes2)
                        RaiseEventSafe(SendProgressEvent, New Object() {Me, send_bytes2})

                        '***** H.H 変更 2013-2-4 *****
                        '***** 送信バッファサイズを超えた分を捨てないように修正 *****
                        If ((total_send_bytes - now_send_bytes) > SendBuffSize) Then
                            will_send_bytes = SendBuffSize
                        Else
                            will_send_bytes = total_send_bytes - now_send_bytes
                        End If

                        Continue Do
                    End If
                End If
                Exit Do
            Loop
            'RaiseEvent SendComplete(Me, send_bytes2)
            RaiseEventSafe(SendCompleteEvent, New Object() {Me, send_bytes2})

        Catch ex As SocketException
            RaiseError("SendDataSub 1", ex.ToString & "(" & ex.ErrorCode & ")")
            CloseSub(0)
        Catch ex As Exception
            RaiseError("SendDataSub 2", ex.ToString)
            CloseSub(0)
        End Try
    End Sub

    Public Sub Close()
        CloseSub(0)
    End Sub

    Private Sub CloseSub(ByVal disc_state As Integer)
        Try
            If (CommSocket IsNot Nothing) Then
                Try
                    CommSocket.Shutdown(SocketShutdown.Both)
                Catch ex As Exception
                End Try
                Try
                    CommSocket.Close()
                Catch ex As Exception
                End Try
            End If
            ChangeState(0)
            'RaiseEvent Closed(Me, 0)
            RaiseEventSafe(ClosedEvent, New Object() {Me, disc_state})
        Catch ex As SocketException
            RaiseError("CloseSub 1", ex.ToString & "(" & ex.ErrorCode & ")")
            ChangeState(0)
            'RaiseEvent Closed(Me, 0)
            RaiseEventSafe(ClosedEvent, New Object() {Me, disc_state})
        Catch ex As Exception
            RaiseError("CloseSub 2", ex.ToString)
            ChangeState(0)
            'RaiseEvent Closed(Me, 0)
            RaiseEventSafe(ClosedEvent, New Object() {Me, disc_state})
        End Try
    End Sub

    Private Sub SockT0001Ctrl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
