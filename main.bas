'simple fem program
'単純な有限要素法プログラムの実現
'平面応力問題か平面ひずみ問題を解く
'解析的な三角形一次要素のみ実装

'お呪い
Option Base 1
Option Explicit

'問題識別コード
Enum ProbType
    ProbNone = 0
    ProbPSS = 1
    ProbPSN = 2
End Enum

'大域データー
Public prob As Integer          '問題番号
Public title As String          'タイトル

'型定義
Public Type FixedDisp
    number As Integer           '番号
    name As String              '名前
    dir(1 To 2) As Integer      '拘束方向
    val(1 To 2) As Double       '拘束値
    data() As Integer           '拘束節点のリスト
End Type

Public Type PointLoad
    number As Integer           '番号
    name As String              '名前
    val(1 To 2) As Double       '荷重値
    data() As Integer           '負荷節点のリスト
End Type

' データーの定義
Public ntnode As Integer        '全節点数
Public ntelem As Integer        '全要素数
Public node() As Integer        '節点番号
Public coord() As Double        '座標値
Public elem() As Integer        '要素番号
Public conn() As Integer        'コネクティビティ

Public area() As Double         '要素の面積
Public A_mat() As Double        'Aマトリックス。Bマトリックスの一つ前の状態という意味

Public young As Double          'ヤング率
Public poisson As Double        'ポアソン比
Public thick As Double          '板厚。平面応力問題では任意。平面ひずみでは常に1
Public D_mat() As Double        'Dマトリックス

Public nfdisp As Integer        '拘束データーの数
Public npload As Integer        '節点荷重データーの数
Public fdisp() As FixedDisp     '拘束データーリスト
Public pload() As PointLoad     '荷重データーリスト

Dim sysK() As Double            '全体剛性マトリックス
Dim copyK() As Double           '全体剛性マトリックスのコピー
Dim rhs() As Double             '方程式を解く際に使われるベクトル
Public dof_list() As Boolean    '拘束条件負荷リスト

Public disp() As Double         '変位ベクトル
Public force() As Double        '荷重ベクトル
Public stress() As Double       '応力
Public strain() As Double       'ひずみ
Public eqiv_force() As Double   '等価節点荷重(結果)

'-------------------------------------------------------------------------------
'ファイル名関連
'取り敢えずファイルの名前は固定する
Private fname As String 'ファイル名
Private Const data_ext As String = "data"   'データー拡張子
Private Const res_a_ext As String = "rsa"   '結果(アスキー)拡張子
Private Const res_b_ext As String = "rsb"   '結果(バイナリ)拡張子
Private Const log_ext As String = "out"     '結果(ログ)拡張子

'読み込んだ行数
Private nline As Integer

'-------------------------------------------------------------------------------
'メインルーチン
Sub main()
    '入出力部分
    open_file
    read_data
    write_res_a
    write_log

    read_bc
    write_bc_res_a
    write_bc_log

    '計算部分
    calc_D_mat
    make_sysK_mat
    set_BC
    solve
    restore_rhs
    calc_result
    calc_eqiv_force

    '結果処理と後始末
    write_result_res_a
    write_result_log
    Close_file
End Sub

'-------------------------------------------------------------------------------
'Dマトリックスを求む
Sub calc_D_mat()
    Dim f As Double, G As Double
    ReDim D_mat(3, 3)

    G = young / 2 / (1 + poisson)
    D_mat(1, 3) = 0: D_mat(2, 3) = 0: D_mat(3, 1) = 0: D_mat(3, 2) = 0
    If prob = ProbPSS Then
        f = young / (1 - poisson * poisson)
        D_mat(1, 1) = f
        D_mat(1, 2) = f * poisson
        D_mat(2, 1) = D_mat(1, 2)
        D_mat(2, 2) = D_mat(1, 1)
        D_mat(3, 3) = G
    ElseIf prob = ProbPSN Then
        f = young / (1 + poisson) / (1 - 2 * poisson)
        D_mat(1, 1) = f * (1 - poisson)
        D_mat(1, 2) = f * poisson
        D_mat(2, 1) = D_mat(1, 2)
        D_mat(2, 2) = D_mat(1, 1)
        D_mat(3, 3) = G
    End If
End Sub

'-------------------------------------------------------------------------------
'行列の積を求む
Function matmul(a() As Double, B() As Double)
    Dim res() As Double
    Dim i As Integer, j As Integer, k As Integer

    ReDim res(UBound(a, 1), UBound(B, 2))
    
    For i = 1 To UBound(a, 1)
        For j = 1 To UBound(B, 2)
            res(i, j) = 0
            For k = 1 To UBound(B, 1)
                res(i, j) = res(i, j) + a(i, k) * B(k, j)
            Next k
        Next j
    Next i
    matmul = res
End Function

'転置行列を求む
Function transpose(a() As Double)
    Dim res() As Double
    Dim i As Integer, j As Integer

    ReDim res(UBound(a, 2), UBound(a, 1))
    
    For i = 1 To UBound(a, 2)
        For j = 1 To UBound(a, 1)
            res(i, j) = a(j, i)
        Next j
    Next i
    transpose = res
End Function

Function cp_mat2(a() As Double)
    Dim res() As Double
    Dim i As Integer, j As Integer

    ReDim res(UBound(a, 1), UBound(a, 2))
    
    For i = 1 To UBound(a, 1)
        For j = 1 To UBound(a, 2)
            res(i, j) = a(i, j)
        Next j
    Next i
    cp_mat2 = res
End Function

'-------------------------------------------------------------------------------
'全体剛性マトリックスを作成する
Sub make_sysK_mat()
    Dim i As Integer, j As Integer, k As Integer
    Dim B_mat(3, 6) As Double, Ke() As Double, r1() As Double, r2() As Double
    Dim nx As Integer, ny As Integer, dx As Integer, dy As Integer
    
    ReDim area(ntelem)
    ReDim A_mat(ntelem, 2, 3)
    ReDim sysK(ntnode * 2, ntnode * 2)
    
    For i = 1 To ntelem
        calc_A_mat i
        make_B_mat B_mat, i
        r1 = transpose(B_mat)
        r2 = matmul(D_mat, B_mat)
        Ke = matmul(r1, r2)
        For j = 1 To 3
            nx = conn(i, j)
            For k = 1 To 3
                ny = conn(i, k)
                sysK(nx * 2 - 1, ny * 2 - 1) = sysK(nx * 2 - 1, ny * 2 - 1) + Ke(j * 2 - 1, k * 2 - 1) * area(i)
                sysK(nx * 2 + 0, ny * 2 - 1) = sysK(nx * 2 + 0, ny * 2 - 1) + Ke(j * 2 + 0, k * 2 - 1) * area(i)
                sysK(nx * 2 - 1, ny * 2 + 0) = sysK(nx * 2 - 1, ny * 2 + 0) + Ke(j * 2 - 1, k * 2 + 0) * area(i)
                sysK(nx * 2 + 0, ny * 2 + 0) = sysK(nx * 2 + 0, ny * 2 + 0) + Ke(j * 2 + 0, k * 2 + 0) * area(i)
            Next k
        Next j
    Next i
End Sub

'要素のAマトリックスを求む
Sub calc_A_mat(i As Integer)
    'ヤコビアン(面積)を求む
    area(i) = (coord(conn(i, 1) * 2 - 1) - coord(conn(i, 3) * 2 - 1)) * (coord(conn(i, 2) * 2) - coord(conn(i, 3) * 2)) _
            - (coord(conn(i, 2) * 2 - 1) - coord(conn(i, 3) * 2 - 1)) * (coord(conn(i, 1) * 2) - coord(conn(i, 3) * 2))
    'Aマトリクスの計算
    A_mat(i, 1, 1) = (coord(conn(i, 2) * 2) - coord(conn(i, 3) * 2)) / area(i)
    A_mat(i, 1, 2) = (coord(conn(i, 3) * 2) - coord(conn(i, 1) * 2)) / area(i)
    A_mat(i, 1, 3) = (coord(conn(i, 1) * 2) - coord(conn(i, 2) * 2)) / area(i)
    A_mat(i, 2, 1) = (coord(conn(i, 3) * 2 - 1) - coord(conn(i, 2) * 2 - 1)) / area(i)
    A_mat(i, 2, 2) = (coord(conn(i, 1) * 2 - 1) - coord(conn(i, 3) * 2 - 1)) / area(i)
    A_mat(i, 2, 3) = (coord(conn(i, 2) * 2 - 1) - coord(conn(i, 1) * 2 - 1)) / area(i)
    area(i) = area(i) / 2
End Sub

'Bマトリックスを求む
Sub make_B_mat(B() As Double, i As Integer)
    Dim j As Integer

    For j = 1 To 3
        B(1, j * 2 - 1) = A_mat(i, 1, j)
        B(1, j * 2 + 0) = 0
        B(2, j * 2 - 1) = 0
        B(2, j * 2 + 0) = A_mat(i, 2, j)
        B(3, j * 2 - 1) = A_mat(i, 2, j)
        B(3, j * 2 + 0) = A_mat(i, 1, j)
    Next j
End Sub

'境界条件の処理
Sub set_BC()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim pos As Integer

    ReDim dof_list(ntnode * 2)
    ReDim rhs(ntnode * 2)
    ReDim disp(ntnode * 2)
    ReDim force(ntnode * 2)
    For i = 1 To nfdisp
        For j = 1 To UBound(fdisp(i).data)
            For k = 1 To 2
                If fdisp(i).dir(k) <> 0 Then
                    pos = (fdisp(i).data(j) - 1) * 2 + k
                    disp(pos) = fdisp(i).val(k)
                    dof_list(pos) = True
                    For l = 1 To ntnode * 2
                        rhs(pos) = rhs(pos) - sysK(l, pos) * disp(pos)
                        sysK(l, pos) = 0
                    Next l
                    sysK(pos, pos) = -1
                End If
            Next k
        Next j
    Next i
    For i = 1 To npload
        For j = 1 To UBound(pload(i).data)
            For k = 1 To 2
                pos = (pload(i).data(j) - 1) * 2 + k
                force(pos) = pload(i).val(k)
                If disp(pos) Then
                    MsgBox "警告:節点荷重テーブル" + pload(i).name + _
                           "内の節点・DOFは拘束されています。荷重は捨てられます。"
                Else
                    rhs(pos) = rhs(pos) + pload(i).val(k)
                End If
            Next k
        Next j
    Next i
End Sub

Sub solve(Optional ByVal tol As Double = 0.0000000001)
    Dim pivot As Double, kij As Double
    Dim i As Integer, j As Integer, k As Integer

    copyK = cp_mat2(sysK)
    '前進消去
    For i = 1 To ntnode * 2
        pivot = copyK(i, i)
        If Abs(pivot) < tol Then
            MsgBox "エラー:" & Str(i) & "行のピボットが許容値未満になりました"
            End
        End If
        For j = i + 1 To ntnode * 2
            kij = copyK(j, i)
            If Abs(kij) >= tol Then
                For k = 1 To ntnode * 2
                    copyK(j, k) = copyK(j, k) - copyK(i, k) / pivot * kij
                Next k
                rhs(j) = rhs(j) - rhs(i) / pivot * kij
            End If
        Next j
    Next i
    '後退代入
    For i = ntnode * 2 To 1 Step -1
        pivot = copyK(i, i)
        rhs(i) = rhs(i) / pivot
        For j = 1 To i - 1
            rhs(j) = rhs(j) - rhs(i) * copyK(j, i)
        Next j
    Next i
End Sub

'求められた結果をそれぞれ変位と荷重に振り分ける。
Sub restore_rhs()
    Dim i As Integer

    For i = 1 To ntnode * 2
        If dof_list(i) Then force(i) = rhs(i) Else disp(i) = rhs(i)
    Next i
End Sub

'結果の計算:応力とひずみを計算する
Sub calc_result()
    Dim i As Integer, j As Integer, k As Integer
    Dim dx As Double, dy As Double
    
    ReDim strain(ntelem, 4)
    ReDim stress(ntelem, 4)
    
    For i = 1 To ntelem
        For j = 1 To 3
            dx = disp((conn(i, j) - 1) * 2 + 1)
            dy = disp((conn(i, j) - 1) * 2 + 2)
            strain(i, 1) = strain(i, 1) + A_mat(i, 1, j) * dx
            strain(i, 2) = strain(i, 2) + A_mat(i, 2, j) * dy
            strain(i, 3) = strain(i, 3) + A_mat(i, 2, j) * dx + A_mat(i, 1, j) * dy
        Next j
        For j = 1 To 3
            For k = 1 To 3
                stress(i, k) = stress(i, k) + D_mat(k, j) * strain(i, j)
            Next k
        Next j
        If prob = ProbPSN Then stress(i, 4) = poisson * (stress(i, 1) + stress(i, 2))
        If prob = ProbPSS Then strain(i, 4) = -poisson / young * (stress(i, 1) + stress(i, 2))
    Next i
End Sub


Sub calc_eqiv_force()
    Dim i As Integer, j As Integer, k As Integer, pos As Integer

    ReDim eqiv_force(ntnode * 2)
    For i = 1 To ntelem
        For j = 1 To 3
            pos = (conn(i, j) - 1) * 2
            eqiv_force(pos + 1) = eqiv_force(pos + 1) + (A_mat(i, 1, j) * stress(i, 1) + A_mat(i, 2, j) * stress(i, 3)) * area(i)
            eqiv_force(pos + 2) = eqiv_force(pos + 2) + (A_mat(i, 2, j) * stress(i, 2) + A_mat(i, 1, j) * stress(i, 3)) * area(i)
        Next j
    Next i
End Sub

'-------------------------------------------------------------------------------
'ファイルのオープン
Sub open_file()
    Dim path As String
    path = "I:\old_data\software\fem"
    fname = "test"
    Open path + "\" + fname + "." + data_ext For Input Access Read As #11
    Open path + "\" + fname + "." + res_a_ext For Output Access Write As #12
    Open path + "\" + fname + "." + log_ext For Output Access Write As #13
End Sub

'ファイルのクローズ
Sub Close_file()
    Close #11
    Close #12
    Close #13
End Sub

'メッシュデーターの読み込み
Sub read_data()
    Dim s As String
    Dim i As Integer, j As Integer, k As Integer
    Dim found As Boolean
    Dim t

    Input #11, title
    nline = nline + 1
    Input #11, prob, ntnode, ntelem, young, poisson, t
    If t = Empty Then thick = 1 Else thick = t '板厚のデフォルト設定
    check_data1
    ReDim node(ntnode)
    ReDim coord(2 * ntnode)
    ReDim conn(ntelem, 3)
    ReDim elem(ntelem)
    For i = 1 To ntnode
        Input #11, node(i), coord(i * 2 - 1), coord(i * 2)
        nline = nline + 1
        For j = i + 1 To ntnode
            If node(i) = node(j) Then
                MsgBox "エラー(" & Str(nline) & "):節点番号" & Str(node(i)) & "が重複しています"
                End
            End If
        Next j
    Next i
    For i = 1 To ntelem
        Input #11, elem(i), conn(i, 1), conn(i, 2), conn(i, 3)
        nline = nline + 1
        For j = i + 1 To ntelem
            If elem(i) = elem(j) Then
                MsgBox "エラー(" & Str(nline) & "):要素番号" & Str(elem(i)) & "が重複しています"
                End
            End If
        Next j
        For j = 1 To 3
            k = search_node_by_number(conn(i, j))
            If k = 0 Then
                MsgBox "エラー(" & Str(nline) & "):要素" & Str(elem(i)) & "の節点" & _
                        Str(node(j)) & "が見つかりませんでした"
                End
            End If
            conn(i, j) = k  '内部節点番号に置き換える。
        Next j
    Next i
End Sub

Function search_node_by_number(ByVal number As Integer) As Integer
    Dim i As Integer
    search_node_by_number = 0
    For i = 1 To ntnode
        If number = node(i) Then
            search_node_by_number = i
            Exit For
        End If
    Next i
End Function

'入力データーチェック1:パラメーターチェック
Sub check_data1()
    If prob <> ProbPSS And prob <> ProbPSN Then
        MsgBox "エラー: 問題番号が正しくありません"
        End
    End If
    If thick <= 0 Then
        MsgBox "エラー: 板厚が0以下になっています。"
        End
    End If
    If young <= 0 Then
        MsgBox "エラー: ヤング率が0以下になっています。"
        End
    End If
    If poisson < 0 Or poisson > 0.5 Then
        MsgBox "エラー: ポアソン比の値が負か0.5を超えています。"
        End
    End If
    If ntnode <= 0 Or ntelem <= 0 Then
        MsgBox "エラー: メッシュデーターがありません。"
        End
    End If
End Sub

'結果ファイルにメッシュデーターを書きだす
Sub write_res_a()
    Dim i As Integer

    Write #12, title
    Write #12, prob, ntnode, ntelem, young, poisson, thick
    For i = 1 To ntnode
        Write #12, node(i), coord(i * 2 - 1), coord(i * 2)
    Next i
    For i = 1 To ntelem
        Write #12, elem(i), conn(i, 1), conn(i, 2), conn(i, 3)
    Next i
End Sub

'ログにメッシュデーターを書きだす
Sub write_log()
    Dim i As Integer
    Dim s As String

    Print #13, "単純なFEMソルバー"
    Print #13, Spc(8); "3角形一次要素の二次元平面応力/ひずみ問題を解く"
    Print #13, ""
    Print #13, "解析タイトル ", title
    
    If prob = ProbPSS Then s = "平面応力" Else s = "平面ひずみ"
    Print #13, "問題 ", s
    Print #13, "物性値"
    Print #13, Spc(8); "ヤング率   ", young
    Print #13, Spc(8); "ポアソン比 ", poisson
    Print #13, "板厚  ", thick
    
    Print #13, "節点リスト"
    Print #13, Spc(8); "節点数", ntnode
    Print #13, Spc(8); "番号", Tab(16); "X座標", Tab(32); "Y座標"
    For i = 1 To ntnode
        Print #13, Tab(8), node(i), coord(i * 2 - 1), coord(i * 2 + 0)
    Next i
    
    Print #13, "要素リスト"
    Print #13, Spc(8); "要素数", ntelem
    Print #13, Spc(8); "番号", Tab(16); "コネクティビティ"
    For i = 1 To ntelem
        Print #13, Tab(8), elem(i), conn(i, 1), conn(i, 2), conn(i, 3)
    Next i
End Sub

'境界条件の読み込み
Sub read_bc()
    Dim i As Integer, j As Integer, k As Integer, node As Integer, ndata As Integer
    Dim found As Boolean

    Input #11, nfdisp, npload
    ReDim fdisp(nfdisp)
    ReDim pload(npload)
    For i = 1 To nfdisp
        Input #11, fdisp(i).number, fdisp(i).dir(1), fdisp(i).dir(2), _
                   fdisp(i).val(1), fdisp(i).val(2), ndata, fdisp(i).name
        nline = nline + 1
        ReDim fdisp(i).data(ndata)
        For j = 1 To ndata
            Input #11, fdisp(i).data(j)
            For k = 1 To ntnode
                node = search_node_by_number(fdisp(i).data(j))
                If node = 0 Then
                    MsgBox "エラー(" & Str(nline) & "):拘束条件テーブル" & Str(i) & "の節点" & _
                            Str(fdisp(i).data(j)) & "が見つかりませんでした"
                    End
                End If
                fdisp(i).data(j) = node
            Next k
        Next j
        nline = nline + 1
    Next i
    For i = 1 To npload
        Input #11, pload(i).number, pload(i).val(1), pload(i).val(2), ndata, pload(i).name
        nline = nline + 1
        ReDim pload(i).data(ndata)
        For j = 1 To ndata
            Input #11, pload(i).data(j)
            For k = 1 To ntnode
                node = search_node_by_number(pload(i).data(j))
                If node = 0 Then
                    MsgBox "エラー(" & Str(nline) & "):節点荷重テーブル" & Str(i) & "の節点" & _
                            Str(pload(i).data(j)) & "が見つかりませんでした"
                    End
                End If
                pload(i).data(j) = node
            Next k
        Next j
        nline = nline + 1
    Next i
End Sub

'境界条件の書き出し
Sub write_bc_res_a()
    Dim i As Integer, j As Integer, ndata As Integer

    For i = 1 To nfdisp
        Write #12, fdisp(i).number, fdisp(i).dir(1), fdisp(i).dir(2), _
                   fdisp(i).val(1), fdisp(i).val(2), fdisp(i).name
        For j = 1 To UBound(fdisp(i).data)
            Write #12, fdisp(i).data(j);
        Next j
    Next i

    For i = 1 To npload
        Write #12, pload(i).number, pload(i).val(1), pload(i).val(2), pload(i).name
        For j = 1 To UBound(pload(i).data)
            Write #12, pload(i).data(j);
        Next j
    Next i
End Sub

Sub write_bc_log()
    Dim i As Integer, j As Integer, ndata As Integer

    Print #13, "境界条件"
    Print #13, ""
    Print #13, "拘束条件"
    For i = 1 To nfdisp
        Print #13, "番号", "方向フラグ", "拘束値リスト", "名前"
        Print #13, fdisp(i).number, fdisp(i).dir(1), fdisp(i).dir(2), _
                   fdisp(i).val(1), fdisp(i).val(2), fdisp(i).name
        Print #13, "拘束節点リスト"
        For j = 1 To UBound(fdisp(i).data)
            Print #13, fdisp(i).data(j);
            If j Mod 8 = 0 Then Print #13, ""
        Next j
        If UBound(fdisp(i).data) Mod 8 <> 0 Then Print #13, ""
    Next i

    Print #13, ""
    Print #13, "節点荷重"
    For i = 1 To npload
        Print #13, "番号", "荷重リスト", "名前"
        Print #13, pload(i).number, pload(i).val(1), pload(i).val(2), pload(i).name
        Print #13, "拘束節点リスト"
        For j = 1 To UBound(pload(i).data)
            Print #13, pload(i).data(j);
            If j Mod 8 = 0 Then Print #13, ""
        Next j
        If UBound(pload(i).data) Mod 8 <> 0 Then Print #13, ""
    Next i
End Sub


Sub write_result_log()
    Dim i As Integer, j As Integer, pos As Integer
    Dim rx As Double, ry As Double

    Print #13, "節点変位"
    Print #13, Spc(8); "番号", "X変位", "Y変位"
    For i = 1 To ntnode
        Print #13, node(i), disp((i - 1) * 2 + 1), disp((i - 1) * 2 + 2)
    Next i
    Print #13, ""

    Print #13, "節点反力"
    Print #13, Spc(8); "番号", "X反力", "Y反力"
    For i = 1 To ntnode
        If dof_list((i - 1) * 2 + 1) Then rx = force((i - 1) * 2 + 1) Else rx = 0
        If dof_list((i - 1) * 2 + 2) Then ry = force((i - 1) * 2 + 2) Else ry = 0
        Print #13, node(i), rx, ry
    Next i
    Print #13, ""

    Print #13, "等価節点荷重"
    Print #13, Spc(8); "番号", "X荷重", "Y荷重"
    For i = 1 To ntnode
        pos = (i - 1) * 2
        Print #13, node(i), eqiv_force(pos + 1), eqiv_force(pos + 2)
    Next i
    Print #13, ""

    Print #13, "ひずみ"
    Print #13, Spc(8); "番号", "Xひずみ", "Yひずみ", "Zひずみ", "XYひずみ"
    For i = 1 To ntelem
        Print #13, elem(i), strain(i, 1), strain(i, 2), strain(i, 4), strain(i, 3)
    Next i
    Print #13, ""

    Print #13, "応力"
    Print #13, Spc(8); "番号", "X応力", "Y応力", "Z応力", "XY応力"
    For i = 1 To ntelem
        Print #13, elem(i), stress(i, 1), stress(i, 2), stress(i, 4), stress(i, 3)
    Next i
    Print #13, ""
End Sub

Sub write_result_res_a()
    Dim i As Integer, j As Integer
    
    For i = 1 To ntnode
        Write #12, node(i), disp((i - 1) * 2 + 1), disp((i - 1) * 2 + 2)
    Next i
    For i = 1 To ntnode
        Write #12, node(i), eqiv_force((i - 1) * 2 + 1), eqiv_force((i - 1) * 2 + 2)
    Next i
    For i = 1 To ntelem
        Write #12, elem(i), strain(i, 1), strain(i, 2), strain(i, 4), strain(i, 3)
    Next i
    For i = 1 To ntelem
        Write #12, elem(i), stress(i, 1), stress(i, 2), stress(i, 4), stress(i, 3)
    Next i
End Sub

