﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.Extension.OpenAI.Tests
{
    public class RestApiServiceTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void CheckRestApiResponse()
        {
            var aiService = new OpenAiService(new CredentialManager());
            var apiKey = aiService.ApiKey;
            var restService = new OpenAiRestApiService(apiKey);

            var requestBody = new
            {
                model = "gpt-4o-mini",
                messages = new[]
            {
                new { role = "system", content = "You are a helpful assistant." },
                new { role = "user", content = "Tell me a joke." }
            },
                max_tokens = 50
            };

            var jsonRequestBody = JsonConvert.SerializeObject(requestBody);


            string result = restService.SendRequest(jsonRequestBody);

            Console.WriteLine(result);

            Assert.That(result, Is.Not.Null);

        }

    }
}
