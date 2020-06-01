#
# Media Dater
# 画像/動画ファイルの撮影日時情報をもとにファイルの名前/日時を変更する
#
# 動作環境:
#   Windows PowerShell 5.1
#   予めスクリプトの実行が許可されている(Set-ExecutionPolicy RemoteSigned)こと
#   このスクリプトファイル(media_dater.ps1)がBOM付UTF-8であること
#
# 対応ファイル:
#   拡張子がjpg/mov/mp4の画像/動画ファイル
#
# 使い方:
#   1. このファイル(media_dater.ps1)を適当な場所へ配置
#   2. 画像/動画ファイルが存在するフォルダ上でPowerShellを開く
#   3. media_dater.ps1 を実行する
#
# 参考:
#   https://neos21.hatenablog.com/entry/2019/11/11/080000
#   https://qiita.com/Kosen-amai/items/52ec7e4e2f15f6a09bc3
#   https://qiita.com/kmr_hryk/items/882b4851e23cec607e70
#
Add-Type -AssemblyName System.Drawing

# Exifから日時文字列を生成する
function getExifDate($path) {
  try {
    $img = New-Object Drawing.Bitmap($path)
  } catch {
    return ""
  }
  $byteAry = ($img.PropertyItems | Where-Object{$_.Id -eq 36867}).Value
  if (!$byteAry) {
    $img.Dispose()
    $img = $null
    return ""
  }

  # "YYYY:MM:DD HH:MM:SS " -> "YYYY/MM/DD HH:MM:SS"
  $byteAry[4] = 47
  $byteAry[7] = 47
  $ret = [System.Text.Encoding]::UTF8.GetString($byteAry)
  $ret = $ret.substring(0, 19)
  $img.Dispose()
  $img = $null

  return $ret
}

# 詳細プロパティから日時文字列を生成する
function getPropDate($folder, $file) {
  $shellFolder = $shellObject.namespace($folder)
  $shellFile = $shellFolder.parseName($file)
  $selectedPropertyNo = ""
  $selectedPropertyName = ""
  $selectedPropertyValue = ""

  for ($i = 0; $i -lt 300; $i++) { # 208まで探せば十分?
    $propertyName = $shellFolder.getDetailsOf($Null, $i)
    if ($propertyName -eq "メディアの作成日時") {
      $propertyValue = $shellFolder.getDetailsOf($shellFile, $i)
      if ($propertyValue) {
        $selectedPropertyNo = $i
        $selectedPropertyName = $propertyName
        $selectedPropertyValue = $propertyValue
        break
      }
    }
  }
  if (!$selectedPropertyNo) {
    return ""
  }

  # " YYYY/ MM/ DD   H:MM" -> "YYYY/MM/DD HH:MM:00"
  $ret = $selectedPropertyValue
  $time = "0" + $ret.substring(16) + ":00" # 秒は00に決め打ち
  $time = $time.substring($time.length - 8, 8)
  $ret = $ret.substring(1, 5) + $ret.substring(7, 3) + $ret.substring(11, 2) + " " + $time

  return $ret
}

# ファイルスキップ時の表示
function printSkipped($fname) {
  Write-Host "$fname (skipped)"
}

# メイン処理
function main {
  Write-Host "== Media Dater v1.0 =="

  # シェルオブジェクトを生成
  $shellObject = New-Object -ComObject Shell.Application

  # ファイルリストを取得
  $targetFiles = Get-ChildItem -File | ForEach-Object { $_.Fullname }

  # ファイル毎の処理
  foreach($targetFile in $targetFiles) {
    $dateStr = ""

    # フォルダパス/ファイル名/拡張子を取得
    $folderPath = Split-Path $targetFile
    $fileName = Split-Path $targetFile -Leaf
    $fileExt = $targetFile.substring($targetFile.length - 3, 3).ToLower() # 3文字固定

    # 日付文字列を取得(YYYY/MM/DD HH:MM:SS)
    if ($fileName.ToLower().endsWith("jpg")) {
      $dateStr = getExifDate $targetFile
    } elseif ($fileName.ToLower().endsWith("mov") -or $fileName.ToLower().endsWith("mp4")) {
      $dateStr = getPropDate $folderPath $fileName
    }
    if (!$dateStr) {
      printSkipped $fileName
      continue
    }

    # ファイル名を変更(YYYYMMDD-HHMMSS-NNN.EXT)
    $renamed = $false
    $newFileName = ""
    $tempFileBase = $dateStr.replace("/", "").replace(" ", "-").replace(":", "")
    for ([int]$i = 0; $i -le 999; $i++)
    {
      $tempPath = $folderPath + "\" + $tempFileBase + "-" + $i.ToString("000") + "." + $fileExt
      # ファイル重複チェック
      if ((Test-Path $tempPath) -eq $false)
      {
        $newFileName = Split-Path $tempPath -Leaf
        try {
          Rename-Item $targetFile -newName $newFileName
        } catch {
          printSkipped $fileName
          continue
        }
        $renamed = $true
        break
      }
    }
    if (!$renamed) {
      printSkipped $fileName
      continue
    }

    # 作成/更新日時を変更
    Set-ItemProperty $newFileName -Name CreationTime -Value $dateStr
    Set-ItemProperty $newFileName -Name LastWriteTime -Value $dateStr

    Write-Host "$fileName -> $newFileName ($dateStr)"
  }

  # シェルオブジェクトを解放
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellObject) | out-null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

# 実行
main
