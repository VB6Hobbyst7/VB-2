<?xml version="1.0" encoding="utf-8"?>
<wsdl:definitions xmlns:soap="http://schemas.xmlsoap.org/wsdl/soap/" xmlns:tm="http://microsoft.com/wsdl/mime/textMatching/" xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:mime="http://schemas.xmlsoap.org/wsdl/mime/" xmlns:tns="http://tempuri.org/" xmlns:s="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://schemas.xmlsoap.org/wsdl/soap12/" xmlns:http="http://schemas.xmlsoap.org/wsdl/http/" targetNamespace="http://tempuri.org/" xmlns:wsdl="http://schemas.xmlsoap.org/wsdl/">
  <wsdl:types>
    <s:schema elementFormDefault="qualified" targetNamespace="http://tempuri.org/">
      <s:element name="UpdateRst">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="sHl7" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="UpdateRstResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="1" maxOccurs="1" name="UpdateRstResult" type="s:int" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="SelectRst">
        <s:complexType />
      </s:element>
      <s:element name="SelectRstResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="SelectRstResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="SelectOrder">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="sHl7" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="SelectOrderResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="SelectOrderResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="SelectMdbOrderList">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="phc_cd" type="s:string" />
            <s:element minOccurs="0" maxOccurs="1" name="sdate" type="s:string" />
            <s:element minOccurs="0" maxOccurs="1" name="edate" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="SelectMdbOrderListResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="SelectMdbOrderListResult">
              <s:complexType>
                <s:sequence>
                  <s:element ref="s:schema" />
                  <s:any />
                </s:sequence>
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="DeleteTestItem">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="sHl7" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="DeleteTestItemResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="DeleteTestItemResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="SelectTestItem">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="sHl7" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="SelectTestItemResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="SelectTestItemResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="InsertTestItem">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="stringdata" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="InsertTestItemResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="InsertTestItemResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="MdbOrderList">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="sHl7" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="MdbOrderListResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="MdbOrderListResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="New_SelectOrder">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="sHl7" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="New_SelectOrderResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="New_SelectOrderResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="User_IDSelect">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="sHl7" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
      <s:element name="User_IDSelectResponse">
        <s:complexType>
          <s:sequence>
            <s:element minOccurs="0" maxOccurs="1" name="User_IDSelectResult" type="s:string" />
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:schema>
  </wsdl:types>
  <wsdl:message name="UpdateRstSoapIn">
    <wsdl:part name="parameters" element="tns:UpdateRst" />
  </wsdl:message>
  <wsdl:message name="UpdateRstSoapOut">
    <wsdl:part name="parameters" element="tns:UpdateRstResponse" />
  </wsdl:message>
  <wsdl:message name="SelectRstSoapIn">
    <wsdl:part name="parameters" element="tns:SelectRst" />
  </wsdl:message>
  <wsdl:message name="SelectRstSoapOut">
    <wsdl:part name="parameters" element="tns:SelectRstResponse" />
  </wsdl:message>
  <wsdl:message name="SelectOrderSoapIn">
    <wsdl:part name="parameters" element="tns:SelectOrder" />
  </wsdl:message>
  <wsdl:message name="SelectOrderSoapOut">
    <wsdl:part name="parameters" element="tns:SelectOrderResponse" />
  </wsdl:message>
  <wsdl:message name="SelectMdbOrderListSoapIn">
    <wsdl:part name="parameters" element="tns:SelectMdbOrderList" />
  </wsdl:message>
  <wsdl:message name="SelectMdbOrderListSoapOut">
    <wsdl:part name="parameters" element="tns:SelectMdbOrderListResponse" />
  </wsdl:message>
  <wsdl:message name="DeleteTestItemSoapIn">
    <wsdl:part name="parameters" element="tns:DeleteTestItem" />
  </wsdl:message>
  <wsdl:message name="DeleteTestItemSoapOut">
    <wsdl:part name="parameters" element="tns:DeleteTestItemResponse" />
  </wsdl:message>
  <wsdl:message name="SelectTestItemSoapIn">
    <wsdl:part name="parameters" element="tns:SelectTestItem" />
  </wsdl:message>
  <wsdl:message name="SelectTestItemSoapOut">
    <wsdl:part name="parameters" element="tns:SelectTestItemResponse" />
  </wsdl:message>
  <wsdl:message name="InsertTestItemSoapIn">
    <wsdl:part name="parameters" element="tns:InsertTestItem" />
  </wsdl:message>
  <wsdl:message name="InsertTestItemSoapOut">
    <wsdl:part name="parameters" element="tns:InsertTestItemResponse" />
  </wsdl:message>
  <wsdl:message name="MdbOrderListSoapIn">
    <wsdl:part name="parameters" element="tns:MdbOrderList" />
  </wsdl:message>
  <wsdl:message name="MdbOrderListSoapOut">
    <wsdl:part name="parameters" element="tns:MdbOrderListResponse" />
  </wsdl:message>
  <wsdl:message name="New_SelectOrderSoapIn">
    <wsdl:part name="parameters" element="tns:New_SelectOrder" />
  </wsdl:message>
  <wsdl:message name="New_SelectOrderSoapOut">
    <wsdl:part name="parameters" element="tns:New_SelectOrderResponse" />
  </wsdl:message>
  <wsdl:message name="User_IDSelectSoapIn">
    <wsdl:part name="parameters" element="tns:User_IDSelect" />
  </wsdl:message>
  <wsdl:message name="User_IDSelectSoapOut">
    <wsdl:part name="parameters" element="tns:User_IDSelectResponse" />
  </wsdl:message>
  <wsdl:portType name="WebServiceSoap">
    <wsdl:operation name="UpdateRst">
      <wsdl:input message="tns:UpdateRstSoapIn" />
      <wsdl:output message="tns:UpdateRstSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="SelectRst">
      <wsdl:input message="tns:SelectRstSoapIn" />
      <wsdl:output message="tns:SelectRstSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="SelectOrder">
      <wsdl:input message="tns:SelectOrderSoapIn" />
      <wsdl:output message="tns:SelectOrderSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="SelectMdbOrderList">
      <wsdl:input message="tns:SelectMdbOrderListSoapIn" />
      <wsdl:output message="tns:SelectMdbOrderListSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="DeleteTestItem">
      <wsdl:input message="tns:DeleteTestItemSoapIn" />
      <wsdl:output message="tns:DeleteTestItemSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="SelectTestItem">
      <wsdl:input message="tns:SelectTestItemSoapIn" />
      <wsdl:output message="tns:SelectTestItemSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="InsertTestItem">
      <wsdl:input message="tns:InsertTestItemSoapIn" />
      <wsdl:output message="tns:InsertTestItemSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="MdbOrderList">
      <wsdl:input message="tns:MdbOrderListSoapIn" />
      <wsdl:output message="tns:MdbOrderListSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="New_SelectOrder">
      <wsdl:input message="tns:New_SelectOrderSoapIn" />
      <wsdl:output message="tns:New_SelectOrderSoapOut" />
    </wsdl:operation>
    <wsdl:operation name="User_IDSelect">
      <wsdl:input message="tns:User_IDSelectSoapIn" />
      <wsdl:output message="tns:User_IDSelectSoapOut" />
    </wsdl:operation>
  </wsdl:portType>
  <wsdl:binding name="WebServiceSoap" type="tns:WebServiceSoap">
    <soap:binding transport="http://schemas.xmlsoap.org/soap/http" />
    <wsdl:operation name="UpdateRst">
      <soap:operation soapAction="http://tempuri.org/UpdateRst" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectRst">
      <soap:operation soapAction="http://tempuri.org/SelectRst" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectOrder">
      <soap:operation soapAction="http://tempuri.org/SelectOrder" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectMdbOrderList">
      <soap:operation soapAction="http://tempuri.org/SelectMdbOrderList" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="DeleteTestItem">
      <soap:operation soapAction="http://tempuri.org/DeleteTestItem" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectTestItem">
      <soap:operation soapAction="http://tempuri.org/SelectTestItem" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="InsertTestItem">
      <soap:operation soapAction="http://tempuri.org/InsertTestItem" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="MdbOrderList">
      <soap:operation soapAction="http://tempuri.org/MdbOrderList" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="New_SelectOrder">
      <soap:operation soapAction="http://tempuri.org/New_SelectOrder" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="User_IDSelect">
      <soap:operation soapAction="http://tempuri.org/User_IDSelect" style="document" />
      <wsdl:input>
        <soap:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
  </wsdl:binding>
  <wsdl:binding name="WebServiceSoap12" type="tns:WebServiceSoap">
    <soap12:binding transport="http://schemas.xmlsoap.org/soap/http" />
    <wsdl:operation name="UpdateRst">
      <soap12:operation soapAction="http://tempuri.org/UpdateRst" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectRst">
      <soap12:operation soapAction="http://tempuri.org/SelectRst" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectOrder">
      <soap12:operation soapAction="http://tempuri.org/SelectOrder" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectMdbOrderList">
      <soap12:operation soapAction="http://tempuri.org/SelectMdbOrderList" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="DeleteTestItem">
      <soap12:operation soapAction="http://tempuri.org/DeleteTestItem" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="SelectTestItem">
      <soap12:operation soapAction="http://tempuri.org/SelectTestItem" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="InsertTestItem">
      <soap12:operation soapAction="http://tempuri.org/InsertTestItem" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="MdbOrderList">
      <soap12:operation soapAction="http://tempuri.org/MdbOrderList" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="New_SelectOrder">
      <soap12:operation soapAction="http://tempuri.org/New_SelectOrder" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
    <wsdl:operation name="User_IDSelect">
      <soap12:operation soapAction="http://tempuri.org/User_IDSelect" style="document" />
      <wsdl:input>
        <soap12:body use="literal" />
      </wsdl:input>
      <wsdl:output>
        <soap12:body use="literal" />
      </wsdl:output>
    </wsdl:operation>
  </wsdl:binding>
  <wsdl:service name="WebService">
    <wsdl:port name="WebServiceSoap" binding="tns:WebServiceSoap">
      <soap:address location="http://10.47.14.52:8009/HL7IFWebService/WebService.asmx" />
    </wsdl:port>
    <wsdl:port name="WebServiceSoap12" binding="tns:WebServiceSoap12">
      <soap12:address location="http://10.47.14.52:8009/HL7IFWebService/WebService.asmx" />
    </wsdl:port>
  </wsdl:service>
</wsdl:definitions>