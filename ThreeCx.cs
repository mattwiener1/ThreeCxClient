using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using MessagePack;

namespace ThreeCx
{

    public class ThreeCx 
    {
        private string? _baseUrl;
        private readonly CookieContainer _cookies = new();

        private readonly HttpClient _httpClient = new();
        public ThreeCx(string baseUrl, string username, string password)
        {
            _baseUrl = baseUrl;
            var url = LoginUrl();
            var auth = new Auth()
            {
                Name = "",
                Username = username,
                Password = password,
            };
            var cookie = new Cookie();
            _cookies.Add(LoginUrl(), cookie);
            var response = _httpClient
                .PostAsJsonAsync(LoginUrl(), auth)
                .Result;
            var authCookies =
                _cookies.GetCookies(LoginUrl());

        }

        private Uri LoginUrl()
        {
            var apiEndPoint = "login";
            _baseUrl = StripHtml(_baseUrl?.Replace("/#", string.Empty).Replace("/login", string.Empty)
                .Replace("//api/", "/api/"));
            _baseUrl = _baseUrl?.Replace("//api/", "/api/");
            var url = new Uri(StripHtml(_baseUrl) + apiEndPoint);
            return url;
        }

        private static string StripHtml(string? input) =>
            Regex.Replace(
                input!,
                "<.*?>",
                string.Empty);
        
    }

    internal class Auth
    {
        public string? Name;
        public string? Username;
        public string? Password;
        

    }
}