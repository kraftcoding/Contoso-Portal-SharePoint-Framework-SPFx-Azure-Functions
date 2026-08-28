class AADGroutestsult {
    [string]$Id
    [bool]$isOffice365Group

    [string] GetTemplateId() { return $this.isOffice365Group ? "c:0o.c|federateddirectoryclaimprovider|$($this.Id)" : "c:0t.c|tenant|$($this.Id)" }


    AADGroutestsult([bool] $isOffice365Group) {
        $this.isOffice365Group = $isOffice365Group
    }
}