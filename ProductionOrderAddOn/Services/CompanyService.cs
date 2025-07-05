using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;

namespace ProductionOrderAddOn.Services
{
    public static class CompanyService
    {
        private static Company _company;

        public static Company GetCompany()
        {
            if (_company == null || !_company.Connected)
            {
                // Get the running DI-API connection from the UI-API
                _company = (Company)Application.SBO_Application.Company.GetDICompany();

                if (!_company.Connected)
                {
                    throw new Exception("Failed to get connected DI-API company object.");
                }
            }
            return _company;
        }
    }
}
