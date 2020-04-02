function Out-BoundView{
    Param(
        [Parameter(ValueFromPipeline = $true)]$viewPath,
        $model
    )
    $view = [String]::Join("",(Get-Content -Path $viewPath))
    $properties = $model.PsObject.Properties | Select-Object -ExpandProperty Name
    foreach($property in $properties){
        $v = $model.($property)
        $view = $view.Replace("{{$($property)}}", $v)
    }
    return $view
}