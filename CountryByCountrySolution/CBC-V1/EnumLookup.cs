using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CountryByCountryReportV1
{
    public static class EnumLookup
    {
        public static OECDDocTypeIndic_EnumType GetOECDDocTypeIndicEnumType(string oECDDocTypeIndic)
        {
            return (OECDDocTypeIndic_EnumType)Enum.Parse(typeof(OECDDocTypeIndic_EnumType), oECDDocTypeIndic);
        }

        public static CountryCode_Type GetCountryCodeEnumType(string countryCodeType)
        {
            return (CountryCode_Type)Enum.Parse(typeof(CountryCode_Type), countryCodeType);
        }

        public static CbcMessageTypeIndic_EnumType GetCbcMessageTypeIndicEnumType(string CbcMessageTypeIndicType)
        {
            return (CbcMessageTypeIndic_EnumType)Enum.Parse(typeof(CbcMessageTypeIndic_EnumType), CbcMessageTypeIndicType);
        }


        public static OECDLegalAddressType_EnumType GetOECDLegalAddressTypeEnumType(string OECDLegalAddressTypeType)
        {
            return (OECDLegalAddressType_EnumType)Enum.Parse(typeof(OECDLegalAddressType_EnumType), OECDLegalAddressTypeType);
        }

        public static CbcReportingRole_EnumType GetCbcReportingRoleEnumType(string CbcReportingRoleType)
        {
            return (CbcReportingRole_EnumType)Enum.Parse(typeof(CbcReportingRole_EnumType), CbcReportingRoleType);
        }



        public static CbcBizActivityType_EnumType GetCbcBizActivityTypeEnumType(string CbcBizActivityTypeEnumType)
        {
            return (CbcBizActivityType_EnumType)Enum.Parse(typeof(CbcBizActivityType_EnumType), CbcBizActivityTypeEnumType);
        }

        //public static UltimateParentEntityRole_EnumType GetUltimateParentEntityRoleEnumType(string UltimateParentEntityRoleEnumType)
        //{
        //    return (UltimateParentEntityRole_EnumType)Enum.Parse(typeof(UltimateParentEntityRole_EnumType), UltimateParentEntityRoleEnumType);
        //}


    }
}
