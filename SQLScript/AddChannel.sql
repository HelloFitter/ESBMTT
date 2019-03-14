INSERT INTO PROTOCOLBIND (PROTOCOLID, BINDTYPE, BINDURI, REQUESTADAPTER, RESPONSEADAPTER, THREADPOOL) VALUES ('&2', 'WSChannelConnector', '<protocol.ws protocolName="WSChannelConnector" id="&2" mode="synchronous" ioDirection="DataIn/DataOut" side="server"><common GroundProtocol="http" ServerType="jetty" wsSecurity="" deployedUDDI="false" serviceids="13000000000" wsType="axis" wsdlURI="http://localhost:Port/EsbWebService/*" isSOAP12="true" autoAssemble="false" /><request Encoding="UTF-8" /><response Encoding="UTF-8" /><advanced threadCount="50" connPerHostCount="200" readTimeout="30000" maxConnCount="2000" soLinger="0" writeBufferSize="2048" reuseAddress="true" readBufferSize="2048" connectionTimeout="30000" tcpNoDelay="true" /></protocol.ws>', 'default_protocolAdapter_req_in', 'default_protocolAdapter_res_in', NULL);
INSERT INTO BINDMAP (SERVICEID, STYPE, LOCATION, VERSION, PROTOCOLID, MAPTYPE) VALUES ('local_in', 'SERVICE', 'local_in', NULL, '&2', 'LOCATION');
INSERT INTO SERVICEINFO (SERVICEID, SERVICETYPE, CONTRIBUTION, PREPARED, GROUPNAME, LOCATION, DESCRIPTION, MODIFYTIME, ADAPTERTYPE, ISCREATE) VALUES ('&1', 'CHANNEL', NULL, NULL, NULL, 'local_in', '&2', '2016-08-11 16:05:52 413', NULL, NULL);
INSERT INTO SERVICES (NAME, INADDRESSID, OUTADDRESSID, TYPE, SESSIONCOUNT, DELIVERYMODE, NODEID, LOCATION, ROUTERABLE) VALUES ('&1', 'a31578d478cb9d20de4c67a8aa75f975', '70022232c41f240289cac75d5b3e3884', 'CHANNEL', 1, '2', NULL, 'local_in', NULL);

INSERT INTO BINDMAP (SERVICEID, STYPE, LOCATION, VERSION, PROTOCOLID, MAPTYPE) VALUES ('&1', 'CHANNEL', 'local_in', NULL, '&2', 'request');
INSERT INTO DATAADAPTER (DATAADAPTERID, DATAADAPTER, LOCATION, ADAPTERTYPE) VALUES ('&1', 'default_comm_soap_channel', 'local_in', 'IN');
COMMIT;

EXIT;
