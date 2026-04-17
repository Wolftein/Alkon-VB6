rename table user_conections to user_connections;

alter table user_connections change CONN_DATE CONNECTION_DATE datetime not null;

alter table user_connections
	add DISCONNECTION_DATE datetime null after CONNECTION_DATE;

alter table user_connections change USER_INDEX ID_USER int(10) null;

alter table user_connections drop column DESCRIP;

drop index IND_NAME on user_connections;

alter table user_connections drop column USER_NAME;



