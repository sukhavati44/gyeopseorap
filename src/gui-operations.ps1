# Loading and worker callbacks must not re-enter project operations through DoEvents.
$script:uiActionDepth=0;$script:uiLoadingDepth=0
$script:operationOpen=${function:NOpen};$script:operationLoad=${function:DLoad}
function NOpen($path){
 $script:uiLoadingDepth++
 try{& $operationOpen $path}finally{$script:uiLoadingDepth=[math]::Max(0,$uiLoadingDepth-1);if(!$busy -and !$uiActionDepth -and !$uiLoadingDepth){$f.UseWaitCursor=$false;Buttons-Enabled $f $true};NRefreshActions}
}
function DLoad($mode,$area){
 $script:uiLoadingDepth++
 try{& $operationLoad $mode $area}finally{$script:uiLoadingDepth=[math]::Max(0,$uiLoadingDepth-1);if(!$busy -and !$uiActionDepth -and !$uiLoadingDepth){$f.UseWaitCursor=$false;Buttons-Enabled $f $true};NRefreshActions}
}
$script:operationSafe=${function:DSafe};$script:operationLegacySafe=${function:Safe}
function DSafe($op){if($busy -or $uiLoadingDepth -gt 0){return};& $operationSafe $op}
function Safe($op){if($busy -or $uiLoadingDepth -gt 0){return};& $operationLegacySafe $op}
$f.Add_FormClosing({param($sender,$eventArgs) if($busy -or $uiLoadingDepth -gt 0){$eventArgs.Cancel=$true}})