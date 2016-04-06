#
# Script.ps1
#

#実行パスを定義
#$FilePath = "C:\Users\miyamam.FAREAST\Downloads\Downloader"
Param([String]$FilePath ="C:\Users\miyamam.FAREAST\Downloads\Downloader")
cd $FilePath

#ファイル名だけを取得（ダウンローダーは取得しない）
#Get-ChildItem C:\Users\miyamam.FAREAST\Downloads\Downloader -Recurse -Name -include *.pptx -Exclude *.exe
foreach($FileName in Get-ChildItem $FilePath -Recurse -Name -include *.pptx -Exclude *.exe){
    #シェルオブジェクトを作成
    $Shell = New-Object -ComObject Shell.Application

    #フォルダの指定
    $Folder = $Shell.NameSpace($FilePath)

    # ファイルの指定
    $File = $Folder.parseName($FileName)
    
    # 詳細プロパティ(撮影日時)の取得
    $GET = $Folder.GetDetailsOf($File,21)
    $SID = $FileName.Substring(0 ,4)

    #タイトルに何も入っていないときはなにもしない
    if($GET -eq ""){
        Write-Host $FileName
    }
    #タイトルにPowerPoint Presentationが入っているときもなにもしない
    elseif($GET.Trim() -eq "PowerPoint Presentation"){
        Write-Host $FileName
    }
    #タイトルに<PowerPoint Presentation>が入っているときもなにもしない
    elseif($GET.Trim() -eq "<Presentation title here>"){
        Write-Host $FileName
    }             
    else{
        
        $Oldfn = $FilePath + '\'+ $FileName
        # \ / ? : * " > < | をクレンジング
        $fn = $GET.Replace(":","")
        $fn = $fn.Replace(" ","_")
        $fn = $fn.Replace("/","")
        $fn = $fn.Replace("?","")
        $fn = $fn.Replace("*","")
        $fn = $fn.Replace(">","")
        $fn = $fn.Replace("<","")
        $fn = $fn.Replace("`"","")

        $Newfn =  $SID + "_" + $fn + ".pptx"
        Move-Item $Oldfn $Newfn -Force

        Write-Host $SID $Newfn
    }

}
