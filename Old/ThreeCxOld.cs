using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using MessagePack;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using Serilog;

#pragma warning disable 618

namespace ThinkVoipTool
{
    public class ThreeCxClient
    {
        private readonly string? _baseUrl;
        private readonly List<RestResponseCookie>? _cookie;
        private readonly string? _passWord;
        private readonly string? _userName;
        private string? _apiEndPoint;
        private RestClient _restClient;
        private RestRequest _restRequest;

        public ThreeCxClient(string? baseUrl, string? username, string? password)
        {
            _userName = username;
            _passWord = password;
            _baseUrl = baseUrl;
            _apiEndPoint = "login";
            _baseUrl = StripHtml(_baseUrl?.Replace("/#", string.Empty).Replace("/login", string.Empty).Replace("//api/", "/api/"));
            _baseUrl = _baseUrl?.Replace("//api/", "/api/");
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddHeader("Content-type", "application/json");
            _restRequest.AddJsonBody(new
            {
                ContentType = "Application/Json",
                Name = "",
                Username = _userName,
                Password = _passWord
            });

            try
            {
                var restResponse = _restClient.Execute(_restRequest);
                _cookie = restResponse.Cookies as List<RestResponseCookie>;
            }
            catch (Exception)
            {
                Logging.Logger.Error($"Failed to connect to 3cx server with the provided info: {_baseUrl}, {_userName}, {_passWord} : e.Message");
                throw;
            }
        }

        internal async Task<string> ResetPassword(string? password)
        {
            _apiEndPoint = "Settings/SecuritySettings";
            var adminPassword = MainWindow.ThreeCxPassword;
            var adminNewPassword = password;
            var adminConfirmNewPassword = password;
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddJsonBody("{\"Param\":\"{}\"}");
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);

            var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Content,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                return "Failed";
            }

            var activeObjectId = properties["Id"].ToString();

            _ = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(activeObjectId, "AdminPassword", adminPassword))
                .ConfigureAwait(false);
            _ = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(activeObjectId, "AdminNewPassword", adminNewPassword))
                .ConfigureAwait(false);
            _ = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(activeObjectId, "AdminConfirmNewPassword", adminConfirmNewPassword))
                .ConfigureAwait(false);

            var saveResult = await SaveUpdate(response, activeObjectId);
            return saveResult;
        }

        public async Task<string> MakeExtensionAdmin(string? extensionNumber)
        {
            var extensionId = await GetExtensionId(extensionNumber).ConfigureAwait(false);
            _apiEndPoint = "ExtensionList/set";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddJsonBody(new
            {
                ContentType = "Application/Json",
                Name = "",
                Id = extensionId
            });
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            if(response.StatusCode.ToString() != "OK")
            {
                throw new Exception();
            }

            var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Content,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                throw new Exception();
            }

            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);


            var id = properties["Id"].ToString();
            await UpdateExtensionAdminSettings(response, id);
            var saveResult = await SaveUpdate(response, id);
            return saveResult;
        }

        public async Task<string?> GetExtensionPinNumber(string? extension)
        {
            var extensionId = await GetExtensionId(extension).ConfigureAwait(false);

            _apiEndPoint = "ExtensionList/set";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddJsonBody(new
            {
                ContentType = "Application/Json",
                Name = "",
                Id = extensionId
            });
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            if(response.StatusCode.ToString() != "OK")
            {
                throw new Exception();
            }

            var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Content,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                throw new Exception();
            }

            var extensionActiveObject = properties["ActiveObject"];
            var extJObject = JObject.Parse(extensionActiveObject.ToString()!);
            var vmPin = extJObject.SelectToken("VMPin");
            if(vmPin == null)
            {
                return "-9999";
            }

            var pinNumber = vmPin["_value"]?.ToString();
            return pinNumber;
        }

        private static string StripHtml(string? input) => Regex.Replace(input!, "<.*?>", string.Empty);


        public async Task<List<Extension>> GetExtensionsList()
        {
            _apiEndPoint = "ExtensionList";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.GET);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var obj = JsonConvert.DeserializeObject<JObject>(response.Content);
                var list = obj.GetValue("list");
                Debug.Assert(list != null, nameof(list) + " != null");
                var results = JsonConvert.DeserializeObject<List<Extension>>(list.ToString());
                return results;
            }
            catch (Exception e)
            {
                Logging.Logger.Error("Failed to find extension list: " + e.Message);
                throw;
            }
        }

        public async Task<List<Phone>> GetPhonesList()
        {
            _apiEndPoint = "PhoneList";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.GET);
            _restRequest.AddHeader("Accept", "application/x-msgpack");
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            var bytes = response.RawBytes;
            List<Phone> results;
            try
            {
                results = MessagePackSerializer.Deserialize<List<Phone>>(bytes);
            }
            catch (Exception)
            {
                results = JsonConvert.DeserializeObject<List<Phone>>(response.Content);
            }

            return results;
        }

        public async Task<ThreeCxLicense> GetThreeCxLicense()
        {
            _apiEndPoint = "License";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.GET);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var license = JsonConvert.DeserializeObject<ThreeCxLicense>(response.Content);

                return license;
            }
            catch (Exception e)
            {
                Logging.Logger.Error("Failed to find 3cx License: " + e.Message);
                throw;
            }
        }

        public async Task<List<InboundRules>> GetThreeCxInboundRules()
        {
            _apiEndPoint = "InboundRulesList";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.GET);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var obj = JsonConvert.DeserializeObject<JObject>(response.Content);
                var list = obj.GetValue("list");
                Debug.Assert(list != null, nameof(list) + " != null");
                var results = JsonConvert.DeserializeObject<List<InboundRules>>(list.ToString());
                return results;
            }
            catch (Exception e)
            {
                Logging.Logger.Error("Failed to find 3cx License: " + e.Message);
                throw;
            }
        }

        public async Task<List<SipTrunk>> GetThreeCxSipTrunks()
        {
            _apiEndPoint = "TrunkList";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.GET);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var obj = JsonConvert.DeserializeObject<JObject>(response.Content);
                var list = obj.GetValue("list");
                Debug.Assert(list != null, nameof(list) + " != null");
                var results = JsonConvert.DeserializeObject<List<SipTrunk>>(list.ToString());
                return results;
            }
            catch (Exception e)
            {
                Log.Error("Failed to find Sip trunks: " + e.Message);
                throw;
            }
        }

        public async Task<Updates> GetUpdatesDay()
        {
            _apiEndPoint = "UpdateChecker/set";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddJsonBody($"{{Username:\"{_userName}\",Password:\"{_passWord}\"}}");
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var result = JsonConvert.DeserializeObject<Updates>(response.Content);
                return result;
            }
            catch (Exception e)
            {
                Logging.Logger.Error("Failed to find update day: " + e.Message);
                throw;
            }
        }

        public async Task<ThreeCxSystemStatus?> GetSystemStatus()
        {
            _apiEndPoint = "SystemStatus";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.GET);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var result = JsonConvert.DeserializeObject<ThreeCxSystemStatus>(response.Content,
                    new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });
                return result;
            }
            catch (Exception e)
            {
                Logging.Logger.Error("Failed to find 3cx system status: " + e.Message);
                throw;
            }
        }

        public async Task<List<Extension>?> GetSystemExtensions()
        {
            _apiEndPoint = "SystemStatus/Extensions";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.GET);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var result = JsonConvert.DeserializeObject<List<Extension>>(response.Content,
                    new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });
                return result;
            }
            catch (Exception e)
            {
                Logging.Logger.Error("Failed to find 3cx system status: " + e.Message);
                throw;
            }
        }


        public async Task<JObject?> GetSipTrunkSettings(string? sipTrunkId)
        {
            _apiEndPoint = "trunklist/set";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            //_restRequest.AddJsonBody($"{{\"Id\":\"{sipTrunkId}\"}}");
            _restRequest.AddJsonBody(new
            {
                ContentType = "Application/Json",
                Name = "",
                Id = sipTrunkId
            });
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            try
            {
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var result = JsonConvert.DeserializeObject<JObject>(response.Content,
                    new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });
                return result;
            }
            catch (Exception e)
            {
                Logging.Logger.Error("Failed to find Sip trunk settings: " + e.Message);
                throw;
            }
        }


        public async Task<string> CreatePhoneOnServer(string? phoneType, string? macAddress, string? extensionNumber)
        {
            var extensionId = await GetExtensionId(extensionNumber).ConfigureAwait(false);


            _apiEndPoint = "ExtensionList/set";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddJsonBody(new
            {
                ContentType = "Application/Json",
                Name = "",
                Id = extensionId
            });
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            if(response.StatusCode.ToString() != "OK")
            {
                throw new Exception();
            }

            var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Content,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                throw new Exception();
            }

            var originalResponse = response;
            var extensionActiveObjectId = properties["Id"].ToString();


            //Create new Phone and get its active object id. 

            _apiEndPoint = "EditorList/new";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);

            var newPhoneProperties = ExtensionPropertyModel.SerializeExtProperty(extensionActiveObjectId, "PhoneDevices", "");
            _restRequest.AddHeader("Content-Type", "application/json;charset=UTF-8");
            _restRequest.AddCookie("CmmSession", originalResponse.Cookies[0].Value!);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            _restRequest.AddHeader("Accept", "*/*");
            _restRequest.AddHeader("Accept-Encoding", "gzip, deflate, br");
            _restRequest.AddHeader("Connection", "keep-alive");
            _restRequest.AddParameter("application/json", newPhoneProperties!, "application/json", ParameterType.RequestBody);
            response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);

            properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Content,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                throw new Exception();
            }

            var phoneActiveObjectId = properties["Id"].ToString();


            //set phone mac and type
            var update = ExtensionPropertyModel.SerializeExtProperty(phoneActiveObjectId, "Model", phoneType);
            var updateResponse = await SendUpdate(response, update);

            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            update = ExtensionPropertyModel.SerializeExtProperty(phoneActiveObjectId, "MacAddress", macAddress);
            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            //save phone to extension
            updateResponse = await SaveUpdate(response, phoneActiveObjectId);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            //save extension to system 
            updateResponse = await SaveUpdate(response, extensionActiveObjectId);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            await UpdatePhoneSettingsOnExtension(extensionId, macAddress, extensionNumber);

            return updateResponse;
        }

        public async Task<List<Phone>> GetListOfPhonesForExtension(string? extensionNumber, string? extensionId)
        {
            _apiEndPoint = "ExtensionList/set";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddJsonBody(new
            {
                ContentType = "Application/Json",
                Name = "",
                Id = extensionId
            });
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);

            if(response.StatusCode.ToString() != "OK")
            {
                throw new Exception();
            }

            var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Content,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                throw new Exception();
            }

            var extensionProperties = JsonConvert.DeserializeObject<Dictionary<string, object>>(properties["ActiveObject"].ToString()!,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            JObject phoneDevices = JsonConvert.DeserializeObject<JObject>(extensionProperties?["PhoneDevices"].ToString()!);
            //var _ = new JArray();
            var phones = JsonConvert.DeserializeObject<JArray>(phoneDevices["_value"]?.ToString()!);
            var phonesList = new List<Phone>();


            foreach (var phone in phones)
            {
                var foundPhone = new Phone();

                var props = JsonConvert.DeserializeObject<Dictionary<string, object>>(phone.ToString(), new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });

                var mac = props?["_str"].ToString();
                var model = props?["Model"].ToString();
                var modelValue = (JObject) JsonConvert.DeserializeObject(model!)!;
                foundPhone.Model = modelValue["_value"]?.ToString();
                foundPhone.ExtensionNumber = extensionNumber;
                foundPhone.MacAddress = mac;

                phonesList.Add(foundPhone);
            }

            return phonesList;
        }

        public async Task<string> UpdatePhoneSettingsOnExtension(string? extensionId, string? macAddress, string? extensionNumber)
        {
            _apiEndPoint = "ExtensionList/set";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddJsonBody(new
            {
                ContentType = "Application/Json",
                Name = "",
                Id = extensionId
            });
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);

            if(response.StatusCode.ToString() != "OK")
            {
                throw new Exception();
            }

            var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(response.Content,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                throw new Exception();
            }

            var extensionActiveObjectId = properties["Id"].ToString();
            var extensionProperties = JsonConvert.DeserializeObject<Dictionary<string, object>>(properties["ActiveObject"].ToString()!,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            JObject phoneDevices = JsonConvert.DeserializeObject<JObject>(extensionProperties?["PhoneDevices"].ToString()!);
            JArray phonesToEdit = JsonConvert.DeserializeObject<JArray>(phoneDevices["_value"]?.ToString()!);
            var idInCollection = "";
            foreach (var phone in phonesToEdit)
            {
                var props = JsonConvert.DeserializeObject<Dictionary<string, object>>(phone.ToString(), new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
                var propId = props?["Id"].ToString();
                var mac = props?["_str"].ToString();
                if(mac?.ToUpper() != macAddress?.ToUpper())
                {
                    continue;
                }

                idInCollection = propId;
            }

            string update = ExtensionExtendedPropertyModel.SerializeExtProperty(extensionActiveObjectId, "PhoneDevices", idInCollection,
                "ScreensaverTimeout", "6 hours");

            var updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            update = ExtensionExtendedPropertyModel.SerializeExtProperty(extensionActiveObjectId, "PhoneDevices", idInCollection,
                "BacklightTimeout", "Always On");

            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            update = ExtensionExtendedPropertyModel.SerializeExtProperty(extensionActiveObjectId, "PhoneDevices", idInCollection,
                "PowerLed", "Voicemails only");
            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            update = ExtensionExtendedPropertyModel.SerializeExtProperty(extensionActiveObjectId, "PhoneDevices", idInCollection,
                "TimeFormat", "12-hour clock (AM/PM)");
            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            update = ExtensionExtendedPropertyModel.SerializeExtProperty(extensionActiveObjectId, "PhoneDevices", idInCollection,
                "ProvisioningMethod", "PROVISIONING_METHOD_STUN");
            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            var extInt = int.Parse(extensionNumber!);
            var localSipPort = 6000 + extInt;

            update = ExtensionExtendedPropertyModel.SerializeExtProperty(extensionActiveObjectId, "PhoneDevices", idInCollection,
                "LocalSipPort", localSipPort);

            try
            {
                updateResponse = await SendUpdate(response, update);
                if(updateResponse == "Failed")
                {
                    return updateResponse;
                }
            }
            catch
            {
                update = ExtensionExtendedPropertyModel.SerializeExtPropertyIntId(extensionActiveObjectId, "PhoneDevices", idInCollection,
                    "LocalSipPort", localSipPort);
                updateResponse = await SendUpdate(response, update);
                if(updateResponse == "Failed")
                {
                    return updateResponse;
                }
            }

            update = ExtensionPropertyModel.SerializeExtProperty(extensionActiveObjectId, "AllowLanOnly", false);
            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            update = ExtensionPropertyModel.SerializeExtProperty(extensionActiveObjectId, "CapabilityReInvite", false);

            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            update = ExtensionPropertyModel.SerializeExtProperty(extensionActiveObjectId, "CapabilityPBXDeliversAudio", true);

            updateResponse = await SendUpdate(response, update);
            if(updateResponse == "Failed")
            {
                return updateResponse;
            }

            updateResponse = await SaveUpdate(response, extensionActiveObjectId);

            return updateResponse;
        }


        public async Task CreateExtensionOnServerFromCsv(string? path)
        {
            const int sharedParksCount = 3;
            using var reader = new StreamReader(path!);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            csv.Context.RegisterClassMap<ExtensionMap>();
            var records = csv.GetRecords<ImportedExtension>();
            foreach (var importedExtension in records)
            {
                var exists = await ExtensionExists(importedExtension.Extension);
                if(!exists)
                {
                    _apiEndPoint = "ExtensionList/new";
                    _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
                    _restRequest = new RestRequest(Method.POST);
                    _restRequest.AddJsonBody("{\"Param\":\"{}\"}");
                }
                else
                {
                    var extensionId = await GetExtensionsList();
                    var extId = extensionId.First(ext => ext.Number == importedExtension.Extension);
                    _apiEndPoint = "ExtensionList/set";
                    _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
                    _restRequest = new RestRequest(Method.POST);
                    _restRequest.AddJsonBody($"{{Id: {extId.Id}}}");
                }


                _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
                var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
                var responseString = response.Content;

                var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(responseString,
                    new JsonSerializerSettings
                    {
                        Formatting = Formatting.Indented,
                        NullValueHandling = NullValueHandling.Ignore
                    });
                if(properties == null)
                {
                    return;
                }

                var id = properties["Id"].ToString();
                var blfConfig = new JObject();
                if(id != null)
                {
                    var cleanProperties = JsonConvert.DeserializeObject<Dictionary<string, object>>(properties["ActiveObject"].ToString()!,
                        new JsonSerializerSettings
                        {
                            Formatting = Formatting.Indented,
                            NullValueHandling = NullValueHandling.Ignore
                        });
                    blfConfig = JsonConvert.DeserializeObject<JObject>(cleanProperties?["BLFConfiguration"].ToString() ??
                                                                       throw new ArgumentNullException());
                }

                var blfOne = blfConfig["_value"];
                if(blfOne == null)
                {
                    return;
                }

                var blfId = JsonConvert.DeserializeObject<JArray>(blfOne.ToString());
                var blfIdList = blfId.Values("Id").ToList();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($@"{importedExtension.Extension}: {importedExtension.Firstname} {importedExtension.Lastname}");
                Console.ResetColor();
                //ext number
                await UpdateExtensionNumber(response, id, importedExtension.Extension);

                //Firstname
                await UpdateExtensionFirstName(response, id, importedExtension.Firstname);

                //Lastname
                await UpdateExtensionLastName(response, id, importedExtension.Lastname);

                //Email address
                await UpdateExtensionEmail(response, id, importedExtension.Email);

                //Mobile Number
                //Check Connectwise for missing info???
                if(importedExtension.MobileNumber != null)
                {
                    await UpdateExtensionMobileNumber(response, id, importedExtension.MobileNumber).ConfigureAwait(false);
                }

                //Caller Id Number
                if(importedExtension.CallerId != null)
                {
                    await UpdateExtensionOutboundCallerId(response, id, importedExtension.CallerId).ConfigureAwait(false);
                }

                //voiceMail options
                if(importedExtension.VoicemailOptions != null)
                {
                    await UpdateExtensionVoiceMailOptions(response, id, importedExtension.VoicemailOptions).ConfigureAwait(false);
                }

                //VoiceMail PIN
                if(importedExtension.Pin != null)
                {
                    await UpdateExtensionVoiceMailPin(response, id, importedExtension.Pin).ConfigureAwait(false);
                }

                //PBX delivers audio
                await UpdateExtensionPbxDeliversAudioOption(response, id).ConfigureAwait(false);

                //Allow use off of LAN
                await UpdateExtensionAllowedUSeOffLan(response, id).ConfigureAwait(false);

                //Disable ReInvites
                await UpdateExtensionDisableReInvites(response, id).ConfigureAwait(false);

                //Accept Multiple Calls
                await UpdateExtensionAcceptMultipleCalls(response, id).ConfigureAwait(false);

                //line Keys
                const int lineKeys = 2;
                await UpdateExtensionLineKeys(lineKeys, response, id, blfIdList).ConfigureAwait(false);

                //Shared parks
                await UpdateExtensionSharedParks(response, id, blfIdList, lineKeys, sharedParksCount).ConfigureAwait(false);
                //Save extension to server

                await SaveExtensionUpdatesToServer(response, id).ConfigureAwait(false);
            }
        }


        public async Task CreateExtensionOnServer(string? extensionNumber, string? firstName, string? lastName, string? email,
            string? voiceMailOptions, string? mobileNumber = "", string? callerId = "", string? pin = "1234", bool disAllowUseOffLan = false,
            bool vmOnly = false, bool fwdOnly = false)
        {
            var exists = await ExtensionExists(extensionNumber).ConfigureAwait(false);
            if(!exists)
            {
                _apiEndPoint = "ExtensionList/new";
                _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
                _restRequest = new RestRequest(Method.POST);
                _restRequest.AddJsonBody(new
                {
                    ContentType = "Application/Json",
                    Name = "",
                    Id = ""
                });
            }
            else
            {
                var currentPin = await GetExtensionPinNumber(extensionNumber);
                if(currentPin != "")
                {
                    pin = currentPin;
                }

                var extId = await GetExtensionId(extensionNumber);
                _apiEndPoint = "ExtensionList/set";
                _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint);
                _restRequest = new RestRequest(Method.POST);
                _restRequest.AddJsonBody(new
                {
                    ContentType = "Application/Json",
                    Name = "",
                    Id = extId
                });
            }

            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            var responseString = response.Content;

            var properties = JsonConvert.DeserializeObject<Dictionary<string, object>>(responseString,
                new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                });
            if(properties == null)
            {
                return;
            }

            var id = properties["Id"].ToString();

            await SendUpdates(extensionNumber, firstName, lastName, email, voiceMailOptions, mobileNumber, callerId, pin, disAllowUseOffLan, vmOnly,
                fwdOnly, response, id).ConfigureAwait(false);
            await SendBlfUpdates(exists, response, properties, id).ConfigureAwait(false);
            await SaveExtensionUpdatesToServer(response, id).ConfigureAwait(false);
        }

        private async Task SendBlfUpdates(bool exists, IRestResponse response, Dictionary<string, object> properties, string? id)
        {
            var blfConfig = new JObject();
            if(id != null)
            {
                var cleanProperties = JsonConvert.DeserializeObject<Dictionary<string, object>>(properties["ActiveObject"].ToString()!,
                    new JsonSerializerSettings
                    {
                        Formatting = Formatting.Indented,
                        NullValueHandling = NullValueHandling.Ignore
                    });
                blfConfig = JsonConvert.DeserializeObject<JObject>(cleanProperties?["BLFConfiguration"].ToString() ??
                                                                   throw new ArgumentNullException());
            }

            var blfOne = blfConfig["_value"];

            if(blfOne != null && !exists)
            {
                var blfId = JsonConvert.DeserializeObject<JArray>(blfOne.ToString());
                var blfIdList = blfId.Values("Id").ToList();

                //line Keys
                const int lineKeys = 2;
                await UpdateExtensionLineKeys(lineKeys, response, id, blfIdList).ConfigureAwait(false);

                //Shared parks
                const int sharedParksCount = 3;
                await UpdateExtensionSharedParks(response, id, blfIdList, lineKeys, sharedParksCount).ConfigureAwait(false);
            }
        }


        private async Task SendUpdates(string? extensionNumber, string? firstName, string? lastName, string? email, string? voiceMailOptions,
            string? mobileNumber, string? callerId, string? pin, bool disAllowUseOffLan, bool vmOnly, bool fwdOnly, IRestResponse response,
            string? id)
        {
            //ext number
            await UpdateExtensionNumber(response, id, extensionNumber).ConfigureAwait(false);

            //Firstname
            await UpdateExtensionFirstName(response, id, firstName).ConfigureAwait(false);

            //Lastname
            await UpdateExtensionLastName(response, id, lastName).ConfigureAwait(false);

            //Mobile Number
            await UpdateExtensionMobileNumber(response, id, mobileNumber).ConfigureAwait(false);

            //Caller Id Number
            await UpdateExtensionOutboundCallerId(response, id, callerId).ConfigureAwait(false);

            if(email != string.Empty)
            {
                //Email address
                await UpdateExtensionEmail(response, id, email).ConfigureAwait(false);
                //voiceMail options
                await UpdateExtensionVoiceMailOptions(response, id, voiceMailOptions).ConfigureAwait(false);
            }

            //VoiceMail PIN
            await UpdateExtensionVoiceMailPin(response, id, pin).ConfigureAwait(false);

            //PBX delivers audio
            await UpdateExtensionPbxDeliversAudioOption(response, id).ConfigureAwait(false);

            //Allow use off of LAN
            await UpdateExtensionAllowedUSeOffLan(response, id, disAllowUseOffLan).ConfigureAwait(false);

            //Disable ReInvites
            await UpdateExtensionDisableReInvites(response, id).ConfigureAwait(false);

            //Accept Multiple Calls
            await UpdateExtensionAcceptMultipleCalls(response, id).ConfigureAwait(false);

            //voicemail Only Extension Forwarding and restriction settings
            if(vmOnly)
            {
                await UpdateForwardingRulesForVmOnly(response, id).ConfigureAwait(false);
                await UpdateRestrictionsForVmOnly(response, id).ConfigureAwait(false);
            }

            if(fwdOnly)
            {
                await UpdateForwardingRulesForFwdOnly(response, id).ConfigureAwait(false);
                await UpdateRestrictionsForFwdOnly(response, id).ConfigureAwait(false);
            }
            else
            {
                await UndoRestrictionsForVmOnly(response, id).ConfigureAwait(false);
                await UndoForwardingRulesForVmOnly(response, id).ConfigureAwait(false);
            }
        }

        private async Task SaveExtensionUpdatesToServer(IRestResponse response, string? id)
        {
            if(await SaveUpdate(response, id) != "OK")
            {
                Logging.Logger.Error($"Failed to save changes for extension Id: {id}  to the server");
                Console.WriteLine(@"Failed to save changes.");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine(@"Create/Update Success");
                Console.ResetColor();
            }
        }

        private async Task UndoRestrictionsForVmOnly(IRestResponse response, string? id)
        {
            var updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "BlockRemoteTunnel", false))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }

            updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AllowWebMeeting", true))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }
        }

        private async Task UndoForwardingRulesForVmOnly(IRestResponse response, string? id)
        {
            var updateResponse = await SendUpdate(response,
                    ExtensionExtendedPropertyModel.SerializeExtFwdProperty(id, "ForwardingAvailable", "NoAnswerTimeout", "20"))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }
        }

        private async Task UpdateRestrictionsForVmOnly(IRestResponse response, string? id)
        {
            var updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "BlockRemoteTunnel", true))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }

            updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AllowWebMeeting", false))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }
        }

        private async Task UpdateRestrictionsForFwdOnly(IRestResponse response, string? id)
        {
            string updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AllowWebMeeting", false))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }
        }

        private async Task UpdateForwardingRulesForVmOnly(IRestResponse response, string? id)
        {
            var updateResponse = await SendUpdate(response,
                    ExtensionExtendedPropertyModel.SerializeExtFwdProperty(id, "ForwardingAvailable", "NoAnswerTimeout", "1"))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }
        }

        private async Task UpdateForwardingRulesForFwdOnly(IRestResponse response, string? id)
        {
            var updateResponse = await SendUpdate(response,
                    ExtensionExtendedPropertyModel.SerializeExtFwdProperty(id, "ForwardingAvailable", "NoAnswerTimeout", "1"))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }
        }


        private async Task UpdateExtensionSharedParks(IRestResponse response, string? id, List<JToken> blfIdList, int lineKeys, int sharedParksCount)
        {
            if(response == null)
            {
                throw new ArgumentNullException(nameof(response));
            }

            if(id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            if(blfIdList == null)
            {
                throw new ArgumentNullException(nameof(blfIdList));
            }

            var blfResponse = await SendBlfUpdate(response,
                    ExtensionExtendedPropertyModel.SerializeExtProperty(id, "BLFConfiguration", blfIdList[2].ToString(), "BlfType",
                        "BlfType.SharedParking"))
                .ConfigureAwait(false);

            for (var key = 2; key < blfIdList.Count; key++)
                blfResponse = await SendBlfUpdate(response,
                    ExtensionExtendedPropertyModel.SerializeExtProperty(id, "BLFConfiguration", blfIdList[key].ToString(), "BlfType",
                        "BlfType.None")).ConfigureAwait(false);

            for (var i = lineKeys; i < sharedParksCount + lineKeys; i++)
            {
                blfResponse = await SendBlfUpdate(response,
                    ExtensionExtendedPropertyModel.SerializeExtProperty(id, "BLFConfiguration", blfIdList[i].ToString(), "BlfType",
                        "BlfType.SharedParking")).ConfigureAwait(false);
                if(blfResponse.StatusCode.ToString() != "OK")
                {
                    Console.WriteLine(@"Failed to Update Blf Options.");
                    throw new Exception();
                }

                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine($@"Blf {i + 1}: ""Blf""  status: OK");
                Console.ResetColor();
            }

            var sharedParkExtensions = JsonConvert.DeserializeObject<JArray>(blfResponse.Content);
            for (var i = 0; i < sharedParksCount; i++)
            {
                if(sharedParkExtensions.Count <= 0)
                {
                    continue;
                }

                var testvalue = sharedParkExtensions[0];
                var values = testvalue["Item"]?["SPExtension"]?["possibleValues"];
                if(values == null)
                {
                    continue;
                }

                var ext = int.Parse(values[i]?["Id"]?.ToString() ?? throw new ArgumentNullException()).ToString();
                var _ = await SendBlfUpdate(response,
                        ExtensionExtendedPropertyModel.SerializeExtProperty(id, "BLFConfiguration", blfIdList[i + 2].ToString(), "SPExtension", ext))
                    .ConfigureAwait(false);
            }
        }

        private async Task UpdateExtensionLineKeys(int lineKeys, IRestResponse response, string? id, List<JToken> blfIdList)
        {
            if(response == null)
            {
                throw new ArgumentNullException(nameof(response));
            }

            if(id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            if(blfIdList == null)
            {
                throw new ArgumentNullException(nameof(blfIdList));
            }

            for (var i = 0; i < lineKeys; i++)
            {
                var responseStatus = await SendUpdate(response,
                        ExtensionExtendedPropertyModel.SerializeExtProperty(id, "BLFConfiguration", blfIdList[i].ToString(), "BlfType",
                            "BlfType.Line"))
                    .ConfigureAwait(false);
                if(responseStatus != "OK")
                {
                    Console.WriteLine(@"Failed to Update Blf Options.");
                    throw new Exception();
                }

                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine($@"Blf {i + 1}: ""Line"" status: OK");
                Console.ResetColor();
            }
        }

        private async Task UpdateExtensionAcceptMultipleCalls(IRestResponse response, string? id)
        {
            var updateResponse = await SendUpdate(response,
                ExtensionPropertyModel.SerializeExtProperty(id, "ForwardingAvailable", "AcceptMultipleCalls", true)).ConfigureAwait(false);
            if(updateResponse != "OK")
            {
                Console.WriteLine(@"Failed to Update Accept Multiple Calls option.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"Accept multiple calls status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionDisableReInvites(IRestResponse response, string? id)
        {
            var responseStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "CapabilityReInvite", false))
                .ConfigureAwait(false);
            if(responseStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update ReInvite option.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"ReInvite options status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionAllowedUSeOffLan(IRestResponse response, string? id)
        {
            var responseStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "AllowLanOnly", false))
                .ConfigureAwait(false);
            if(responseStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update use off of LAN option.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"Use extension off LAN options status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionAllowedUSeOffLan(IRestResponse response, string? id, bool enable)
        {
            var responseStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "AllowLanOnly", enable))
                .ConfigureAwait(false);
            if(responseStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update use off of LAN option.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"Use extension off LAN options status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionPbxDeliversAudioOption(IRestResponse response, string? id)
        {
            var responseStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "CapabilityPBXDeliversAudio", true))
                .ConfigureAwait(false);
            if(responseStatus != "OK")
            {
                Console.WriteLine(@"PBX delivers audio options.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"PBX delivers audio status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionVoiceMailPin(IRestResponse response, string? id, string? pin)
        {
            var enabledResponse =
                await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "VMEnabled", true)).ConfigureAwait(false);
            if(enabledResponse == "Failed")
            {
                throw new Exception();
            }

            var updateResponse = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "VMPin", pin)).ConfigureAwait(false);
            if(updateResponse != "OK")
            {
                Console.WriteLine(@"Failed to Update VoiceMail PIN.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"VoiceMail PIN status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionVoiceMailOptions(IRestResponse response, string? id, string? voiceMailOptions)
        {
            var enabledResponse =
                await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "VMEnabled", true)).ConfigureAwait(false);
            if(enabledResponse == "Failed")
            {
                throw new Exception();
            }

            var responseStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "VMEmailOptions", voiceMailOptions))
                .ConfigureAwait(false);

            if(responseStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update VoiceMail Email Options.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"VoiceMail email options status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionOutboundCallerId(IRestResponse response, string? id, string? callerId)
        {
            var updateStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "OutboundCallerId", callerId))
                .ConfigureAwait(false);

            if(updateStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update caller Id property.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"Caller Id status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionMobileNumber(IRestResponse response, string? id, string? mobileNumber)
        {
            var updateStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "MobileNumber", mobileNumber))
                .ConfigureAwait(false);
            if(updateStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update mobile phone number property.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"Mobile phone status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionEmail(IRestResponse response, string? id, string? email)
        {
            var updateStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "Email", email)).ConfigureAwait(false);
            if(updateStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update Email property.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"Email status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionLastName(IRestResponse response, string? id, string? lastName)
        {
            var updateStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "LastName", lastName))
                .ConfigureAwait(false);

            if(updateStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update LastName property.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"LastName status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionFirstName(IRestResponse response, string? id, string? firstName)
        {
            var updateStatus = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "FirstName", firstName))
                .ConfigureAwait(false);
            if(updateStatus != "OK")
            {
                Console.WriteLine(@"Failed to Update FirstName property.");
                throw new Exception();
            }

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine(@"FirstName status: OK");
            Console.ResetColor();
        }

        private async Task UpdateExtensionNumber(IRestResponse response, string? id, string? extensionNumber)
        {
            var exists = await ExtensionExists(extensionNumber);
            if(!exists)
            {
                var updateResponse = await SendUpdate(response, ExtensionPropertyModel.SerializeExtProperty(id, "Number", extensionNumber))
                    .ConfigureAwait(false);
                if(updateResponse != "OK")
                {
                    Console.WriteLine(@"Failed to Update Extension Number property.");
                    throw new Exception();
                }

                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine(@"Extension number status: OK");
                Console.ResetColor();
            }
        }

        private async Task UpdateExtensionAdminSettings(IRestResponse response, string? id)
        {
            var updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AccessEnabled", true))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }

            updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AccessRole", "AccessRole.GlobalExtensionManager"))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }

            updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AccessAdmin", true))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }


            updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AccessReporter", true))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }

            updateResponse = await SendUpdate(response,
                    ExtensionPropertyModel.SerializeExtProperty(id, "AccessReporterRecording", true))
                .ConfigureAwait(false);
            if(updateResponse == "Failed")
            {
                throw new Exception();
            }
        }

        private async Task<string> SendUpdate(IRestResponse originalResponse, string? update)
        {
            _apiEndPoint = "edit/update";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint) {Timeout = -1};
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddHeader("Content-Type", "application/json;charset=UTF-8");
            _restRequest.AddCookie("CmmSession", originalResponse.Cookies[0].Value!);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            _restRequest.AddHeader("Accept", "*/*");
            _restRequest.AddHeader("Accept-Encoding", "gzip, deflate, br");
            _restRequest.AddHeader("Connection", "keep-alive");
            _restRequest.AddParameter("application/json", update!, "application/json", ParameterType.RequestBody);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            return response.StatusCode == HttpStatusCode.OK ? response.StatusCode.ToString() : "Failed";
        }

        public async Task<string?> GetExtensionId(string? extensionNUmber)
        {
            var extensionId = await GetExtensionsList().ConfigureAwait(false);
            return extensionId.First(ext => ext.Number == extensionNUmber).Id;
        }

        public async Task DeleteExtension(string? extensionNumber)
        {
            var iD = await GetExtensionId(extensionNumber).ConfigureAwait(false);
            _apiEndPoint = "ExtensionList/delete";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint) {Timeout = -1};
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddHeader("Content-Type", "application/json;charset=UTF-8");
            //_restRequest.AddCookie("CmmSession", originalResponse.Cookies[0].Value!);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            _restRequest.AddHeader("Accept", "*/*");
            _restRequest.AddHeader("Accept-Encoding", "gzip, deflate, br");
            _restRequest.AddHeader("Connection", "keep-alive");
            _restRequest.AddParameter("application/json", $"{{Ids: [{iD}]}}", "application/json", ParameterType.RequestBody);
            _ = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
        }

        private async Task<IRestResponse> SendBlfUpdate(IRestResponse originalResponse, string? update)
        {
            _apiEndPoint = "edit/update";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint) {Timeout = -1};
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddHeader("Content-Type", "application/json;charset=UTF-8");
            _restRequest.AddCookie("CmmSession", originalResponse.Cookies[0].Value!);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            _restRequest.AddHeader("Accept", "*/*");
            _restRequest.AddHeader("Accept-Encoding", "gzip, deflate, br");
            _restRequest.AddHeader("Connection", "keep-alive");
            _restRequest.AddParameter("application/json", update!, "application/json", ParameterType.RequestBody);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            return response;
        }

        private async Task<string> SaveUpdate(IRestResponse originalResponse, string? update)
        {
            _apiEndPoint = "edit/save";
            _restClient = new RestClient(StripHtml(_baseUrl) + _apiEndPoint) {Timeout = -1};
            _restRequest = new RestRequest(Method.POST);
            _restRequest.AddHeader("Content-Type", "application/json;charset=UTF-8");
            _restRequest.AddCookie("CmmSession", originalResponse.Cookies[0].Value!);
            _restRequest.AddCookie(_cookie?[0].Name!, _cookie?[0].Value!);
            _restRequest.AddHeader("Accept", "*/*");
            _restRequest.AddHeader("Accept-Encoding", "gzip, deflate, br");
            _restRequest.AddHeader("Connection", "keep-alive");
            _restRequest.AddParameter("application/json", update!, "application/json", ParameterType.RequestBody);
            var response = await _restClient.ExecuteAsync(_restRequest).ConfigureAwait(false);
            return response.StatusCode != HttpStatusCode.OK ? "Failed" : response.StatusCode.ToString();
        }

        private async Task<bool> ExtensionExists(string? extensionNumber)
        {
            var extensionsList = await GetExtensionsList().ConfigureAwait(false);

            return extensionsList.Count(ext => ext.Number == extensionNumber) > 0;
        }

        private sealed class ExtensionMap : ClassMap<ImportedExtension>
        {
            private ExtensionMap()
            {
                Map(m => m.Extension).Name("Extension");
                Map(m => m.Firstname).Name("Firstname");
                Map(m => m.Lastname).Name("Lastname");
                Map(m => m.Email).Name("Email");
            }
        }
    }

    internal static class ExtensionPropertyModel
    {
        public static string SerializePhoneProperty(string? objectId, string? propertyPath) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\"}}]}},\"Param\":\"{{}}\"}}";

        public static string SerializeExtProperty(string? objectId, string? propertyPath, string? value) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\"}}]}},\"PropertyValue\":\"{value}\"}}";

        public static string SerializeExtProperty(string? objectId, string? propertyPath, string? propertyPathTwo, bool value) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\"}},{{\"Name\":\"{propertyPathTwo}\"}}]}},\"PropertyValue\":{value.ToString().ToLower()}}}";

        public static string SerializeExtProperty(string? objectId, string? propertyPath, bool value) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\"}}]}},\"PropertyValue\":{value.ToString().ToLower()}}}";
    }

    internal static class ExtensionExtendedPropertyModel
    {
        public static string SerializeExtProperty(string? objectId, string? propertyPath, string? idInCollection, string? name,
            string? propertyValue) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\",\"IdInCollection\":\"{idInCollection}\"}},{{\"Name\":\"{name}\"}}]}},\"PropertyValue\":\"{propertyValue}\"}}";

        public static string SerializeExtProperty(string? objectId, string? propertyPath, string? idInCollection, string? name, int propertyValue) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\",\"IdInCollection\":\"{idInCollection}\"}},{{\"Name\":\"{name}\"}}]}},\"PropertyValue\":{propertyValue}}}";

        public static string SerializeExtPropertyIntId(string? objectId, string? propertyPath, string? idInCollection, string? name,
            int propertyValue) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\",\"IdInCollection\":{idInCollection}}},{{\"Name\":\"{name}\"}}]}},\"PropertyValue\":{propertyValue}}}";


        public static string SerializeExtFwdProperty(string? objectId, string? propertyPath, string? propertyPath2, string? propertyValue) =>
            $"{{\"Path\":{{\"ObjectId\":\"{objectId}\",\"PropertyPath\":[{{\"Name\":\"{propertyPath}\"}},{{\"name\":\"{propertyPath2}\"}}]}},\"PropertyValue\":{propertyValue}}}";
    }

    public class ImportedExtension
    {
        public ImportedExtension(string? extension, string? firstname, string? lastname, string? email, string? mobileNumber, string? callerId,
            string? voicemailOptions, string? pin)
        {
            Extension = extension;
            Firstname = firstname;
            Lastname = lastname;
            Email = email;
            MobileNumber = mobileNumber;
            CallerId = callerId;
            VoicemailOptions = voicemailOptions;
            Pin = pin;
        }

        public string? Extension { get; }
        public string? Firstname { get; }
        public string? Lastname { get; }
        public string? Email { get; }
        public string? MobileNumber { get; }
        public string? CallerId { get; }
        public string? VoicemailOptions { get; }
        public string? Pin { get; }
    }

    public class Extension
    {
        public string? Id { get; set; }

        //public bool IsOperator { get; set; }
        public bool IsRegistered { get; set; }


        public string? Number { get; set; }

        public string? FirstName { get; set; }
        public string? LastName { get; set; }

        public string? Email { get; set; }

        //public string? Password { get; set; }
        public string? MobileNumber { get; set; }

        //public string? OutboundCallerId { get; set; }
        public int Phones { get; set; }

        public string? MacAddress { get; set; }
        // public string? Membership { get; set; }
        // public string? CurrentProfile { get; set; }
        // public int QueueStatus { get; set; }

        //[JsonProperty("DND")]
        //public int Dnd { get; set; }

        //public Warning Warning { get; set; }
    }

    public class ThreeCxServer
    {
        public ThreeCxLicense? ThreeCxLicense { get; private set; }
        public List<Phone>? Phones { get; private set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        // ReSharper disable once UnusedAutoPropertyAccessor.Local
        private List<Extension>? Extensions { get; set; }
        public List<InboundRules>? InboundRulesList { get; private set; }
        public List<SipTrunk>? SipTrunks { get; private set; }
        public string? UpdateDay { get; private set; }
        public ThreeCxSystemStatus? SystemStatus { get; private set; }
        public JObject? SipTrunkSettings { get; private set; }

        public async Task<ThreeCxServer> Create(ThreeCxClient client)
        {
            var server = new ThreeCxServer();
            await server.ThreeCxServerInitialize(client);
            return server;
        }

        private async Task ThreeCxServerInitialize(ThreeCxClient client)
        {
            Phones = await client.GetPhonesList();
            Extensions = await client.GetExtensionsList();
            ThreeCxLicense = await client.GetThreeCxLicense();
            InboundRulesList = await client.GetThreeCxInboundRules();
            SipTrunks = await client.GetThreeCxSipTrunks();
            var updates = await client.GetUpdatesDay();
            UpdateDay = updates.ActiveObject.ScheduleDay.Selected?.Replace("DayOfWeek.", string.Empty);
            SystemStatus = await client.GetSystemStatus();
            SipTrunkSettings = await client.GetSipTrunkSettings(SipTrunks[0].Id);
        }
    }

    public class ThreeCxSystemStatus
    {
        [JsonProperty("FQDN")]
        public string? Fqdn { get; set; }

        [JsonProperty("WebMeetingFQDN")]
        public string? WebMeetingFqdn { get; set; }

        [JsonProperty("WebMeetingBestMCU")]
        public string? WebMeetingBestMcu { get; set; }

        public string? Version { get; set; }
        public int RecordingState { get; set; }
        public bool Activated { get; set; }
        public int MaxSimCalls { get; set; }
        public int MaxSimMeetingParticipants { get; set; }
        public long CallHistoryCount { get; set; }
        public int ChatMessagesCount { get; set; }
        public int ExtensionsRegistered { get; set; }
        public bool OwnPush { get; set; }

        // ReSharper disable once UnassignedGetOnlyAutoProperty
        public string? Ip { get; }
        public bool LocalIpValid { get; set; }
        public string? CurrentLocalIp { get; set; }
        public string? AvailableLocalIps { get; set; }
        public int ExtensionsTotal { get; set; }
        public bool HasUnregisteredSystemExtensions { get; set; }
        public bool HasNotRunningServices { get; set; }
        public int TrunksRegistered { get; set; }
        public int TrunksTotal { get; set; }
        public int CallsActive { get; set; }
        public int BlacklistedIpCount { get; set; }
        public int MemoryUsage { get; set; }
        public int PhysicalMemoryUsage { get; set; }
        public long FreeVirtualMemory { get; set; }
        public long TotalVirtualMemory { get; set; }
        public long FreePhysicalMemory { get; set; }
        public long TotalPhysicalMemory { get; set; }
        public int DiskUsage { get; set; }
        public long FreeDiskSpace { get; set; }
        public long TotalDiskSpace { get; set; }
        public long CpuUsage { get; set; }
        public List<List<object>>? CpuUsageHistory { get; set; }
        public string? MaintenanceExpiresAt { get; set; }
        public bool Support { get; set; }
        public string? ExpirationDate { get; set; }
        public int OutboundRules { get; set; }
        public bool BackupScheduled { get; set; }
        public object? LastBackupDateTime { get; set; }
        public string? ResellerName { get; set; }
        public string? LicenseKey { get; set; }
        public string? ProductCode { get; set; }
        public bool IsSpla { get; set; }
    }

    public class SipTrunk
    {
        public SipTrunk(string? externalNumber, string? simCalls, string? type, string? host, string? name, string? id)
        {
            ExternalNumber = externalNumber;
            SimCalls = simCalls;
            Type = type;
            Host = host;
            Name = name;
            Id = id;
        }

        public string? Id { get; }

        //public string? Str { get; set; }
        //public string? Number { get; set; }
        public string? Name { get; }
        public string? Host { get; }
        public string? Type { get; }
        public string? SimCalls { get; }
        public string? ExternalNumber { get; }

        //public bool IsRegistered { get; set; }
        //public Gateway Gateway { get; set; }
        //public AuthId authID { get; set; }

        //public class AuthId
        //{
        //    public string? type { get; set; }
        //    public string? _value { get; set; }
        //    public bool disabled { get; set; }
        //}

        //public class AuthPassword
        //{
        //    public string? type { get; set; }
        //    public string? _value { get; set; }
        //}

        //public class MainDidNumber
        //{
        //    public string? type { get; set; }
        //    public string? _value { get; set; }
        //}

        //public class SimultaneousCalls
        //{
        //    public string? type { get; set; }
        //    public string? _value { get; set; }
        //}
    }

    //public class Gateway
    //{
    //    public bool disabled { get; set; }
    //    public _value _Value { get; set; }
    //}

    //public class _value
    //{
    //    public string? Id { get; set; }
    //    public string? Str { get; set; }
    //    public TypeOfGateway typeOfGateway { get; set; }

    //    public class TypeOfGateway
    //    {
    //        public string? selected { get; set; }
    //    }
    //}

    public class InboundRules
    {
        public string? Name { get; set; }
        public string? Trunk { get; set; }
        public string? Did { get; set; }
        public string? InOfficeRouting { get; set; }
        public string? OutOfOfficeRouting { get; set; }


        //public class InOfficeRoute
        //{
        //    public string? Type { get; set; }
        //    public string? Dn { get; set; }
        //    public string? Voicemail { get; set; }
        //    public string? ExternalNumber { get; set; }
        //}

        //public class OutOfOfficeRoute
        //{
        //    public string? Type { get; set; }
        //    public string? Dn { get; set; }
        //    public string? Voicemail { get; set; }
        //    public string? ExternalNumber { get; set; }
        //}
    }

    public class ThreeCxLicense
    {
        public ThreeCxLicense(string? key, int maxSimCalls)
        {
            Key = key;
            MaxSimCalls = maxSimCalls;
        }

        public string? Key { get; }
        public string? CompanyName { get; set; }
        public string? ContactName { get; set; }
        public string? Email { get; set; }
        public string? AdminEMail { get; set; }
        public string? Telephone { get; set; }
        public string? ResellerName { get; set; }
        public string? ProductCode { get; set; }
        public int MaxSimCalls { get; }
        public bool ProFeatures { get; set; }
        public string? ExpirationDate { get; set; }
    }

    [MessagePackObject]
    public class Phone
    {
        private string? _modelShortName;


        [Key(0)]
        public string? WhateverThisThingIs { get; set; }

        [Key(1)]
        public int Id { get; set; }

        [IgnoreMember]
        public string? UserAgent { get; set; }

        [Key(3)]
        public string? LastRegistration { get; set; }

        [IgnoreMember]
        public string? ProvMethod { get; set; }

        [IgnoreMember]

        public int DeviceType { get; set; }

        [Key(6)]
        public string? Model { get; set; }

        [IgnoreMember]
        public string? ModelShortName { get; set; }

        [IgnoreMember]
        public string? ModelDisplayName
        {
            get => _modelShortName ?? Model;
            set => _modelShortName = value;
        }

        [Key(7)]
        public string? Vendor { get; set; }

        [Key(8)]
        public string? FirmwareVersion { get; set; }

        [Key(9)]
        public string? Name { get; set; }

        [Key(10)]
        public string? UserId { get; set; }

        [Key(11)]
        public string? UserPassword { get; set; }

        [Key(12)]
        public string? Pin { get; set; }

        [Key(13)]
        public string? Ip { get; set; }

        [Key(14)]
        public string? InterfaceLink { get; set; }

        [IgnoreMember]

        public int SipPort { get; set; }

        [Key(16)]
        public string? MacAddress { get; set; }

        [IgnoreMember]

        public string? Status { get; set; }

        [Key(18)]
        public string? PhoneWebPassword { get; set; }

        [Key(19)]
        public string? ProvLink { get; set; }

        [Key(20)]
        public bool IsNew { get; set; }

        [Key(21)]
        public bool AssignExtensionEnabled { get; set; }

        [Key(22)]
        public bool CanBeRebooted { get; set; }

        [Key(23)]
        public bool CanBeUpgraded { get; set; }

        [Key(244)]
        public bool CanBeProvisioned { get; set; }

        [Key(25)]
        public bool HasInterface { get; set; }

        [Key(26)]
        public bool IsCustomProvisionTemplate { get; set; }

        [Key(27)]
        public bool UnsupportedFirmware { get; set; }

        [Key(28)]
        public string? HotdeskingExtension { get; set; }

        [IgnoreMember]

        public string? DisplayText { get; set; }

        [Key(30)]
        public string? ExtensionNumber { get; set; }
    }


    public class Updates
    {
        public Updates(ActiveObject activeObject)
        {
            ActiveObject = activeObject;
        }

        public int Id { get; set; }
        public ActiveObject ActiveObject { get; }
    }

    public class ReadyToSave
    {
        [JsonProperty("type")]
        public string? Type { get; set; }

        [JsonProperty("hide")]
        public bool Hide { get; set; }

        [JsonProperty("_value")]
        public string? Value { get; set; }
    }

    public class ScheduleType
    {
        [JsonProperty("type")]
        public string? Type { get; set; }

        [JsonProperty("selected")]
        public string? Selected { get; set; }

        [JsonProperty("possibleValues")]
        public List<string>? PossibleValues { get; set; }

        [JsonProperty("translatable")]
        public bool Translatable { get; set; }
    }

    public class ScheduleDay
    {
        [JsonProperty("type")]
        public string? Type { get; set; }

        [JsonProperty("selected")]
        public string? Selected { get; set; }

        [JsonProperty("possibleValues")]
        public List<string>? PossibleValues { get; set; }

        [JsonProperty("translatable")]
        public bool Translatable { get; set; }
    }

    public class ScheduleTime
    {
        [JsonProperty("type")]
        public string? Type { get; set; }

        [JsonProperty("_value")]
        public string? Value { get; set; }
    }

    public class TcxPbxUpdates
    {
        [JsonProperty("type")]
        public string? Type { get; set; }

        [JsonProperty("_value")]
        public string? Value { get; set; }
    }

    public class ActiveObject
    {
        public ActiveObject(ScheduleDay scheduleDay)
        {
            ScheduleDay = scheduleDay;
        }

        public string? Id { get; set; }

        [JsonProperty("_str")]
        public string? Str { get; set; }

        public bool IsNew { get; set; }
        public ReadyToSave? ReadyToSave { get; set; }
        public ScheduleType? ScheduleType { get; set; }
        public ScheduleDay ScheduleDay { get; }
        public ScheduleTime? ScheduleTime { get; set; }
        public TcxPbxUpdates? TcxPbxUpdates { get; set; }
    }

    public class Warning
    {
        [JsonProperty("invalidPasswords")]
        public List<string>? InvalidPasswords { get; set; }

        [JsonProperty("highAlert")]
        public string? HighAlert { get; set; }
    }
}