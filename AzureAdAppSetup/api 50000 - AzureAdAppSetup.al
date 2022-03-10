page 50000 AzureAdAppSetup
{
    PageType = API;
    Caption = 'AzureAdAppSetup';
    APIPublisher = 'Microsoft';
    APIGroup = 'Setup';
    APIVersion = 'beta', 'v1.0';
    EntityName = 'aadApp';
    EntitySetName = 'aadApps';
    SourceTable = "Name/Value Buffer";
    DelayedInsert = true;
    SourceTableTemporary = true;

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field(id; "ID")
                {
                    Caption = 'ID';
                }
                field(name; "Name")
                {
                    Caption = 'Name';
                }
                field(value; "Value")
                {
                    Caption = 'Value';
                }
            }
        }
    }

    trigger OnInsertRecord(BelowxRec: Boolean): Boolean
    begin
        case Name of
            'SetupAzureAdApp':
                begin
                    SetupAzureAdApp(Value);
                end;
        end;
    end;

    local procedure SetupAzureAdApp(argument: Text)
    var
        AzureAdAppSetup: Record "Azure AD App Setup";
        AzureADMgt: Codeunit "Azure AD Mgt.";
    begin
        if not AzureAdAppSetup.FindFirst then
            AzureAdAppSetup.Init;

        AzureAdAppSetup."Redirect URL" := AzureADMgt.GetRedirectUrl;
        AzureAdAppSetup."App ID" := SelectStr(1, argument);
        AzureAdAppSetup.SetSecretKeyToIsolatedStorage(SelectStr(2, argument));

        if not AzureAdAppSetup.Modify(true) then
            AzureAdAppSetup.Insert(true);
    end;

}