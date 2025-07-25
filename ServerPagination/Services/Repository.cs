using Microsoft.Extensions.Options;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using Newtonsoft.Json;
using ServerPagination.Models.Comman;
using RestSharp;
namespace ServerPagination.Services
{
    public class ManageService<TRequestEntity, TResponseEntity> : IManageService<TRequestEntity, TResponseEntity>
    {
        private readonly AppSettings appSettings;
        public ManageService(IOptions<AppSettings> _appSettings)
        {
            appSettings = _appSettings.Value;
        }
        public async Task<TResponseEntity> GetAsync(string aRequestUri, int id)
        {
            try
            {
                var client = new RestClient(appSettings.APIUri);
                var request = new RestRequest(aRequestUri + id, Method.Get);
                request.AddHeader("Content-Type", "application/json");

                var vSvcResponse = await client.ExecuteAsync(request);
                if (vSvcResponse.IsSuccessful)
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    return response.data;
                }
                else
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    throw new InvalidOperationException(response.message);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public async Task<List<TResponseEntity>> GetAllByIdAsync(string aRequestUri, string id)
        {
            try
            {
                var client = new RestClient(appSettings.APIUri);
                var request = new RestRequest(aRequestUri + id, Method.Get);
                request.AddHeader("Content-Type", "application/json");

                var vSvcResponse = await client.ExecuteAsync(request);
                if (vSvcResponse.IsSuccessful)
                {
                    var response = JsonConvert.DeserializeObject<ListResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    return response.data;
                }
                else
                {
                    var response = JsonConvert.DeserializeObject<ListResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    throw new InvalidOperationException(response.message);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public async Task<List<TResponseEntity>> GetAllAsync(string aRequestUri)
        {
            try
            {
                var client = new RestClient(appSettings.APIUri);
                var request = new RestRequest(aRequestUri, Method.Get);
                request.AddHeader("Content-Type", "application/json");

                var vSvcResponse = await client.ExecuteAsync(request);

                if (vSvcResponse.IsSuccessful)
                {
                    var response = JsonConvert.DeserializeObject<ListResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    return response.data;
                }
                else
                {
                    var response = JsonConvert.DeserializeObject<ListResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    throw new InvalidOperationException(response.message);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public async Task<TResponseEntity> PostAsync(string aRequestUri, TRequestEntity aObj)
        {
            try
            {
                var client = new RestClient(appSettings.APIUri);
                var request = new RestRequest(aRequestUri, Method.Post);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", JsonConvert.SerializeObject(aObj), ParameterType.RequestBody);

                var vSvcResponse = await client.ExecuteAsync(request);

                if (vSvcResponse.IsSuccessful)
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    return response.data;
                }
                else
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    throw new InvalidOperationException(response.message);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public async Task<TResponseEntity> PutAsync(string aRequestUri, TRequestEntity aObj)
        {
            try
            {
                var client = new RestClient(appSettings.APIUri);
                var request = new RestRequest(aRequestUri, Method.Put);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", JsonConvert.SerializeObject(aObj), ParameterType.RequestBody);


                var vSvcResponse = await client.ExecuteAsync(request);
                if (vSvcResponse.IsSuccessful)
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    return response.data;
                }
                else
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    throw new InvalidOperationException(response.message);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public async Task<TResponseEntity> DeleteAsync(string aRequestUri, int id)
        {
            try
            {
                var client = new RestClient(appSettings.APIUri);
                var request = new RestRequest(aRequestUri + id, Method.Delete);
                request.AddHeader("Content-Type", "application/json");

                var vSvcResponse = await client.ExecuteAsync(request);
                if (vSvcResponse.IsSuccessful)
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    return response.data;
                }
                else
                {
                    var response = JsonConvert.DeserializeObject<ResponseModel<TResponseEntity>>(vSvcResponse.Content);
                    throw new InvalidOperationException(response.message);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}