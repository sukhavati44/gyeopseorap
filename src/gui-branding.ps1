$iconPath=Join-Path $base 'assets/gyeopseorap-open.ico'
$imagePath=Join-Path $base 'assets/gyeopseorap-open.png'
if(Test-Path -LiteralPath $iconPath){$f.Icon=New-Object Drawing.Icon($iconPath)}
if(Test-Path -LiteralPath $imagePath){
 $sourceImage=[Drawing.Image]::FromFile($imagePath);try{$script:productImage=$sourceImage.Clone()}finally{$sourceImage.Dispose()}
 $homeLogo=New-Object Windows.Forms.PictureBox;$homeLogo.Image=$productImage;$homeLogo.SizeMode='Zoom';$homeLogo.SetBounds(24,10,44,44);$homeLogo.TabStop=$false;$homeLogo.AccessibleName='겹서랍 열린 서랍 아이콘';$nHome.Controls.Add($homeLogo);$homeLogo.BringToFront()
 $workLogo=New-Object Windows.Forms.PictureBox;$workLogo.Image=$productImage;$workLogo.SizeMode='Zoom';$workLogo.SetBounds(12,4,28,28);$workLogo.TabStop=$false;$workLogo.AccessibleName='겹서랍 열린 서랍 아이콘';$dashboard.Controls.Add($workLogo);$workLogo.BringToFront()
 $f.Add_FormClosed({$productImage.Dispose()})
}
