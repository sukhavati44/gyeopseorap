function Manage-Membership($req){
 if(!$M.students.ContainsKey($Student)){throw '등록된 학생을 선택하세요.'};$start=[datetime]::MinValue;if(![datetime]::TryParseExact([string]$req.from,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$start)){throw '적용일은 yyyy-MM-dd 형식'}
 $className=([string]$req.class).Trim();if(!$className){throw '학급명을 입력하세요.'};$current=@($M.memberships|Where-Object {$_.student -eq $Student -and (!$_.until -or $_.until -ge $req.from)})
 if(@($current|Where-Object {$_.from -ge $req.from}).Count){throw '같은 날 또는 이후 소속 이력이 있습니다. 기존 이력을 자동 교체하지 않습니다.'}
 Copy-Item $MetaPath (Join-Path $Root ('backups/membership-'+[guid]::NewGuid().ToString('N')+'.json'))
 foreach($entry in $current){$entry.until=$start.AddDays(-1).ToString('yyyy-MM-dd')}
 $M.memberships+=@{student=$Student;class=$className;academic_year=$start.Year;from=[string]$req.from;until=$null;created_at=(Get-Date -Format o);reason=[string]$req.note};Save-Meta
 return @{student_id=$Student;memberships=@($M.memberships|Where-Object student -eq $Student);roster_identity_unchanged=$true}
}
