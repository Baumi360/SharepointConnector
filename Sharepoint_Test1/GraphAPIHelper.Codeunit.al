codeunit 50113 "Graph API Helper"
{
    var
        OAuth2: Codeunit OAuth2;
        ClientIdTxt: Label '0b042988-9423-4865-b204-c4a144a9a08e', Locked = true;
        ClientSecret: Label '9dfb2e0b-b617-4a34-aeae-8ea47fb463f3', Locked = true;
        ResourceUrlTxt: Label 'https://graph.microsoft.com', Locked = true;
        OAuthAuthorityUrlTxt: Label 'https://login.microsoftonline.com/5e553341-24f0-4e09-895e-cd8cc0521d1e/oauth2/v2.0/authorize', Locked = true;
        RedirectURLTxt: Label 'https://localhost:8080/login', Locked = true;
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