codeunit 50113 "Graph API Helper"
{
    var
        OAuth2: Codeunit OAuth2;
        ClientIdTxt: Label 'bcad1f4c-0f2f-4d70-85f6-5355b3a21e9d', Locked = true;
        ClientSecret: Label 'gRS7Q~Zv5GUmvSj58O_ZHfkaj.4EOuJCHELZZ', Locked = true;
        ResourceUrlTxt: Label 'https://graph.microsoft.com', Locked = true;
        OAuthAuthorityUrlTxt: Label 'https://login.microsoftonline.com/5e553341-24f0-4e09-895e-cd8cc0521d1e/oauth2/authorize', Locked = true;
        RedirectURLTxt: Label 'https://businesscentral.dynamics.com/OAuthLanding.htm', Locked = true;
        OneDriveRootQueryUri: Label 'https://graph.microsoft.com/v1.0/me/drive/root/children', Locked = true;

    procedure GetAccessToken(): Text
    var
        PromptInteraction: Enum "Prompt Interaction";
        AccessToken: Text;
        AuthCodeError: Text;
    begin

        OAuth2.AcquireTokenByAuthorizationCode(
            ClientIdTxt,
            ClientSecret,
            OAuthAuthorityUrlTxt,
            RedirectURLTxt,
            ResourceURLTxt,
            PromptInteraction::Consent,
            AccessToken,
            AuthCodeError);

        if (AccessToken = '') or (AuthCodeError <> '') then
            Error(AuthCodeError);

        exit(AccessToken);
    end;



    /// <summary>
    /// GetOneDriveFiles.
    /// </summary>
    /// <returns>Return value of type JsonObject.</returns>
    procedure GetOneDriveFiles(): JsonObject
    var
        Client: HttpClient;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        JsonResponse: JsonObject;
        AccessToken: Text;
        JsonContent: Text;
    begin
        AccessToken := GetAccessToken();
        Dialog.Message(AccessToken);
        RequestMessage.Method('GET');
        RequestMessage.SetRequestUri(OneDriveRootQueryUri);
        Client.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));
        Client.DefaultRequestHeaders().Add('Accept', 'application/json');

        if Client.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.HttpStatusCode() = 200 then begin
                ResponseMessage.Content.ReadAs(JsonContent);
                JsonResponse.ReadFrom(JsonContent);
                exit(JsonResponse);
            end;
    end;
}