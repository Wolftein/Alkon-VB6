--
-- Table structure for table `account_bank`
--

DROP TABLE IF EXISTS `ACCOUNT_BANK`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `ACCOUNT_BANK` (
  `ID_ACCOUNT` int(10) NOT NULL DEFAULT 0,
  `PASSWORD` varchar(32) DEFAULT NULL,
  `SLOT` tinyint(4) NOT NULL DEFAULT 0,
  `OBJ_INDEX` mediumint(9) DEFAULT NULL,
  `AMOUNT` mediumint(9) DEFAULT NULL,
  PRIMARY KEY (`ID_ACCOUNT`,`SLOT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `account_char_share`
--

DROP TABLE IF EXISTS `ACCOUNT_CHAR_SHARE`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `ACCOUNT_CHAR_SHARE` (
  `ID_ACCOUNT_OWNER` int(10) DEFAULT NULL,
  `ID_ACCOUNT_SHARED` int(10) NOT NULL DEFAULT 0,
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `SHARE_DATE` date DEFAULT NULL,
  `SHARE_TIME` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`ID_ACCOUNT_SHARED`,`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;


--
-- Table structure for table `account_info`
--

DROP TABLE IF EXISTS `ACCOUNT_INFO`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `ACCOUNT_INFO` (
  `ID_ACCOUNT` int(10) NOT NULL AUTO_INCREMENT,
  `NAME` varchar(30) DEFAULT NULL,
  `EMAIL` varchar(50) DEFAULT NULL,
  `PASSWORD` varchar(32) DEFAULT NULL,
  `SECRET_QUESTION` varchar(50) DEFAULT NULL,
  `ANSWER` varchar(50) DEFAULT NULL,
  `ACTIVATION_CODE` varchar(32) DEFAULT NULL,
  `STATUS` tinyint(4) DEFAULT NULL,
  `BAN_DETAIL` varchar(100) DEFAULT NULL,
  `CREATION_DATE` datetime DEFAULT NULL,
  `BANK_GOLD` bigint(20) DEFAULT 0,
  `BANK_PASSWORD` varchar(32) DEFAULT NULL,
  PRIMARY KEY (`ID_ACCOUNT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `codes`
--

DROP TABLE IF EXISTS `CODES`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `CODES` (
  `DESCRIP` varchar(20) NOT NULL,
  `CODE` smallint(6) DEFAULT NULL,
  PRIMARY KEY (`DESCRIP`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `critic_events`
--

DROP TABLE IF EXISTS `CRITIC_EVENTS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `CRITIC_EVENTS` (
  `ID_EVENT` int(10) NOT NULL AUTO_INCREMENT,
  `EVENT_DATE` datetime DEFAULT NULL,
  `DESCRIP` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`ID_EVENT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `errores`
--

DROP TABLE IF EXISTS `ERRORES`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `ERRORES` (
  `ID_ERROR` int(10) NOT NULL AUTO_INCREMENT,
  `EVENT_DATE` date DEFAULT NULL,
  `EVENT_TIME` varchar(20) DEFAULT NULL,
  `DESCRIP` varchar(10000) DEFAULT NULL,
  PRIMARY KEY (`ID_ERROR`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Table structure for table `nicknames_blacklist`
--

DROP TABLE IF EXISTS `NICKNAMES_BLACKLIST`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `NICKNAMES_BLACKLIST` (
  `NAME` varchar(30) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `premium_emails`
--

DROP TABLE IF EXISTS `PREMIUM_EMAILS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `PREMIUM_EMAILS` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `EMAIL` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `punishment_basetype`
--

DROP TABLE IF EXISTS `PUNISHMENT_BASETYPE`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `PUNISHMENT_BASETYPE` (
  `ID` tinyint(4) NOT NULL AUTO_INCREMENT,
  `DESCRIPTION` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `punishment_type`
--

DROP TABLE IF EXISTS `PUNISHMENT_TYPE`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `PUNISHMENT_TYPE` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `DESCRIPTION` varchar(50) DEFAULT NULL,
  `BASE_TYPE` tinyint(6) DEFAULT NULL,
  `ENABLED` bit(1) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `punishment_type_rules`
--

DROP TABLE IF EXISTS `PUNISHMENT_TYPE_RULES`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `PUNISHMENT_TYPE_RULES` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ID_PUNISHMENT_TYPE` int(11) DEFAULT NULL,
  `PUNISHMENT_COUNT` smallint(6) NOT NULL COMMENT 'En este campo se guarda la cantidad de penas de un tipo determinado para que se aplique la regla.',
  `PUNISHMENT_SEVERITY` int(255) DEFAULT NULL COMMENT 'Cantidad de dias / horas que se aplican a esta pena',
  `ADD_BAN` bit(1) DEFAULT b'0',
  `ADD_JAIL` bit(1) DEFAULT b'0',
  `NEXT_PUNISHMENT_ID` int(11) DEFAULT -1,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `statictics`
--

DROP TABLE IF EXISTS `STATICTICS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `STATICTICS` (
  `ID_STATICTIC` int(10) NOT NULL AUTO_INCREMENT,
  `EVENT_DATE` datetime DEFAULT NULL,
  `ID_USER` int(10) DEFAULT NULL,
  `DESCRIP` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`ID_STATICTIC`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `store_items`
--

DROP TABLE IF EXISTS `STORE_ITEMS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `STORE_ITEMS` (
  `ID` int(20) NOT NULL AUTO_INCREMENT,
  `ID_CHAR` int(20) NOT NULL,
  `ID_ITEM_WEBSITE` int(20) NOT NULL,
  `ITEM_NUMBER` varchar(45) NOT NULL,
  `ITEM_SLOT` varchar(45) NOT NULL,
  `AMOUNT` int(20) NOT NULL,
  `UNIT_PRICE` int(20) NOT NULL,
  `FROM_INVENTORY` bit(1) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `temp_memory`
--

DROP TABLE IF EXISTS `TEMP_MEMORY`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `TEMP_MEMORY` (
  `slot` int(11) NOT NULL,
  PRIMARY KEY (`slot`)
) ENGINE=MEMORY DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_attributes`
--

DROP TABLE IF EXISTS `USER_ATTRIBUTES`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_ATTRIBUTES` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `STRENGHT` tinyint(4) DEFAULT NULL,
  `DEXERITY` tinyint(4) DEFAULT NULL,
  `INTELLIGENCE` tinyint(4) DEFAULT NULL,
  `CHARISM` tinyint(4) DEFAULT NULL,
  `HEALTH` tinyint(4) DEFAULT NULL,
  PRIMARY KEY (`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_ban_detail`
--

DROP TABLE IF EXISTS `USER_BAN_DETAIL`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_BAN_DETAIL` (
  `ID_BAN` int(10) NOT NULL AUTO_INCREMENT,
  `EVENT_DATE` datetime DEFAULT NULL,
  `ID_USER` int(10) DEFAULT NULL,
  `BANED_BY` varchar(30) DEFAULT NULL,
  `DETAIL` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`ID_BAN`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_bank`
--

DROP TABLE IF EXISTS `USER_BANK`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_BANK` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `SLOT` tinyint(4) NOT NULL DEFAULT 0,
  `OBJ_INDEX` mediumint(9) DEFAULT NULL,
  `AMOUNT` mediumint(9) DEFAULT NULL,
  PRIMARY KEY (`ID_USER`,`SLOT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_conditions`
--

DROP TABLE IF EXISTS `user_conditions`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE IF NOT EXISTS `user_conditions` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `ID_USER` int(10) unsigned DEFAULT NULL,
  `PERSISTED_AT` datetime NOT NULL,
  `CONDITION_PAYLOAD` text NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Table structure for table `user_conections`
--

DROP TABLE IF EXISTS `USER_CONNECTIONS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_CONNECTIONS` (
  `ID_EVENT` int(20) NOT NULL AUTO_INCREMENT,
  `ID_USER` int(10) DEFAULT NULL,
  `CONNECTION_DATE` datetime DEFAULT NULL,
  `DISCONNECTION_DATE` datetime DEFAULT NULL,
  `IP` varchar(15) DEFAULT NULL,
  PRIMARY KEY (`ID_EVENT`),
  KEY `IND_USERID` (`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_faction`
--

DROP TABLE IF EXISTS `USER_FACTION`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_FACTION` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `ALIGNMENT` TINYINT(4) UNSIGNED NULL DEFAULT NULL,
  `ARMY` tinyint(4) DEFAULT NULL,
  `CHAOS` tinyint(4) DEFAULT NULL,
  `NEUTRAL_KILLED` mediumint(9) DEFAULT NULL,
  `CITY_KILLED` mediumint(9) DEFAULT NULL,
  `CRI_KILLED` mediumint(9) DEFAULT NULL,
  `CHAOS_ARMOUR_GIVEN` tinyint(4) DEFAULT NULL,
  `ARMY_ARMOUR_GIVEN` tinyint(4) DEFAULT NULL,
  `CHAOS_EXP_GIVEN` tinyint(4) DEFAULT NULL,
  `ARMY_EXP_GIVEN` tinyint(4) DEFAULT NULL,
  `CHAOS_REWARD_GIVEN` tinyint(4) DEFAULT NULL,
  `ARMY_REWARD_GIVEN` tinyint(4) DEFAULT NULL,
  `NUM_SIGNS` int(11) DEFAULT NULL,
  `SIGNING_LEVEL` tinyint(4) DEFAULT NULL,
  `SIGNING_DATE` datetime DEFAULT NULL,
  `SIGNING_KILLED` mediumint(9) DEFAULT NULL,
  `NEXT_REWARD` int(11) DEFAULT NULL,
  `ROYAL_COUNCIL` tinyint(4) DEFAULT NULL,
  `CHAOS_COUNCIL` tinyint(4) DEFAULT NULL,
  `EXPELLER` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_flags`
--

DROP TABLE IF EXISTS `USER_FLAGS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_FLAGS` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `MUERTO` tinyint(4) DEFAULT NULL,
  `ESCONDIDO` tinyint(4) DEFAULT NULL,
  `HAMBRE` tinyint(3) unsigned DEFAULT NULL,
  `SED` tinyint(3) unsigned DEFAULT NULL,
  `DESNUDO` tinyint(4) DEFAULT NULL,
  `BAN` tinyint(4) DEFAULT NULL,
  `NAVEGANDO` tinyint(4) DEFAULT NULL,
  `ENVENENADO` tinyint(4) DEFAULT NULL,
  `PARALIZADO` tinyint(4) DEFAULT NULL,
  `LAST_MAP` smallint(6) DEFAULT NULL,
  `LAST_TAMED_PET` mediumint(9) DEFAULT 0,
  PRIMARY KEY (`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_info`
--

DROP TABLE IF EXISTS `USER_INFO`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_INFO` (
  `ID_USER` int(10) NOT NULL AUTO_INCREMENT,
  `NAME` varchar(30) DEFAULT NULL,
  `GENDER` tinyint(4) DEFAULT NULL,
  `RACE` tinyint(4) DEFAULT NULL,
  `CLASS` tinyint(4) DEFAULT NULL,
  `HOME` tinyint(4) DEFAULT NULL,
  `DESCRIP` varchar(50) DEFAULT NULL,
  `PUNISHMENT` tinyint(4) DEFAULT NULL,
  `HEADING` tinyint(4) DEFAULT NULL,
  `HEAD` mediumint(9) DEFAULT NULL,
  `BODY` mediumint(9) DEFAULT NULL,
  `WEAPON_ANIM` mediumint(9) DEFAULT NULL,
  `SHIELD_ANIM` mediumint(9) DEFAULT NULL,
  `HELMET_ANIM` mediumint(9) DEFAULT NULL,
  `UP_TIME` int(11) DEFAULT NULL,
  `LAST_IP` varchar(20) DEFAULT NULL,
  `LAST_POS` varchar(11) DEFAULT NULL,
  `WEAPON_SLOT` tinyint(4) DEFAULT NULL,
  `ARMOUR_SLOT` tinyint(4) DEFAULT NULL,
  `HELMET_SLOT` tinyint(4) DEFAULT NULL,
  `SHIELD_SLOT` tinyint(4) DEFAULT NULL,
  `BOAT_SLOT` tinyint(4) DEFAULT NULL,
  `MUNITION_SLOT` tinyint(4) DEFAULT NULL,
  `SACKPACK_SLOT` tinyint(4) DEFAULT NULL,
  `RING_SLOT` tinyint(4) DEFAULT NULL,
  `GUILD_ID` int(11) DEFAULT NULL,
  `REQUESTING_GUILD` mediumint(9) DEFAULT NULL,
  `GUILD_REJECT_DETAIL` varchar(100) DEFAULT NULL,
  `BANNED` tinyint(4) DEFAULT NULL,
  `ID_BAN_PUNISHMENT` int(11) DEFAULT NULL,
  `LOGGED` tinyint(4) DEFAULT NULL,
  `TRAINNING_TIME` int(11) DEFAULT NULL,
  `ID_ACCOUNT` int(10) DEFAULT NULL,
  `PRIVILEGE` tinyint(4) DEFAULT 0,
  `CREATION_DATE` datetime DEFAULT NULL,
  PRIMARY KEY (`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_inventory`
--

DROP TABLE IF EXISTS `USER_INVENTORY`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_INVENTORY` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `SLOT` tinyint(4) NOT NULL DEFAULT 0,
  `OBJ_INDEX` mediumint(9) DEFAULT NULL,
  `AMOUNT` mediumint(9) DEFAULT NULL,
  `EQUIPPED` tinyint(4) DEFAULT NULL,
  PRIMARY KEY (`ID_USER`,`SLOT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_messages`
--

DROP TABLE IF EXISTS `USER_MESSAGES`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_MESSAGES` (
  `ID_USER` int(10) DEFAULT NULL,
  `MSG_INDEX` smallint(6) DEFAULT NULL,
  `MESSAGE` varchar(100) DEFAULT NULL,
  `UNREAD` smallint(6) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_pets`
--

DROP TABLE IF EXISTS `USER_PETS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_PETS` (
  `ID_USER` int(10) DEFAULT NULL,
  `NUM_PET` tinyint(4) DEFAULT NULL,
  `NPC_INDEX` mediumint(9) DEFAULT NULL,
  `NPC_TYPE` mediumint(9) DEFAULT NULL,
  `NPC_LIFE` mediumint(9) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_polls`
--

DROP TABLE IF EXISTS `USER_POLLS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_POLLS` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `POLL_ID` mediumint(9) NOT NULL DEFAULT 0,
  `POLL_OPTION` tinyint(4) DEFAULT NULL,
  `EMAIL` varchar(50) DEFAULT NULL,
  `EVENT_DATE` datetime DEFAULT NULL,
  PRIMARY KEY (`ID_USER`,`POLL_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_punishment`
--

DROP TABLE IF EXISTS `USER_PUNISHMENT`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_PUNISHMENT` (
  `ID_PUNISHMENT` int(10) NOT NULL AUTO_INCREMENT,
  `ID_USER` int(10) DEFAULT NULL,
  `ID_PUNISHER` int(10) DEFAULT NULL,
  `ID_PUNISHMENT_TYPE` int(10) DEFAULT NULL,
  `EVENT_DATE` datetime DEFAULT NULL,
  `END_DATE` datetime DEFAULT NULL,
  `REASON` varchar(255) DEFAULT NULL,
  `ADMIN_NOTES` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`ID_PUNISHMENT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_skills`
--

DROP TABLE IF EXISTS `USER_SKILLS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_SKILLS` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `SKILL` tinyint(4) NOT NULL DEFAULT 0,
  `NATURAL_AMOUNT` tinyint(4) DEFAULT NULL,
  `ASSIGNED_AMOUNT` tinyint(4) DEFAULT NULL,
  `SKILL_EXP_NEXT_LEVEL` mediumint(9) DEFAULT NULL,
  `SKILL_EXP` mediumint(9) DEFAULT NULL,
  PRIMARY KEY (`ID_USER`,`SKILL`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_spells`
--

DROP TABLE IF EXISTS `USER_SPELLS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_SPELLS` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `SLOT` tinyint(4) NOT NULL DEFAULT 0,
  `SPELL_INDEX` mediumint(9) DEFAULT NULL,
  PRIMARY KEY (`ID_USER`,`SLOT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `user_stats`
--

DROP TABLE IF EXISTS `USER_STATS`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `USER_STATS` (
  `ID_USER` int(10) NOT NULL DEFAULT 0,
  `ORO` int(11) DEFAULT NULL,
  `ORO_BANCO` int(11) DEFAULT NULL,
  `HP_MAX` smallint(6) DEFAULT NULL,
  `HP_MIN` smallint(6) DEFAULT NULL,
  `STAMINA_MAX` smallint(6) DEFAULT NULL,
  `STAMINA_MIN` smallint(6) DEFAULT NULL,
  `MANA_MAX` smallint(6) DEFAULT NULL,
  `MANA_MIN` smallint(6) DEFAULT NULL,
  `HIT_MAX` smallint(6) DEFAULT NULL,
  `HIT_MIN` smallint(6) DEFAULT NULL,
  `AGUA_MAX` smallint(6) DEFAULT NULL,
  `AGUA_MIN` smallint(6) DEFAULT NULL,
  `HAMBRE_MAX` smallint(6) DEFAULT NULL,
  `HAMBRE_MIN` smallint(6) DEFAULT NULL,
  `SKILLS` smallint(6) DEFAULT NULL,
  `EXP` bigint(20) DEFAULT NULL,
  `NIVEL` tinyint(4) DEFAULT NULL,
  `EXP_NEXT` int(11) DEFAULT NULL,
  `USERS_KILLED` mediumint(9) DEFAULT NULL,
  `NPCS_KILLED` mediumint(9) DEFAULT NULL,
  `RANKING_POINTS` int(10) UNSIGNED DEFAULT 0,
  `MASTERY_POINTS` int(11) DEFAULT 0,
  `DUELOS_GANADOS` mediumint(6) DEFAULT 0,
  `DUELOS_PERDIDOS` mediumint(6) DEFAULT 0,
  `ORO_DUELOS` int(11) DEFAULT 0,
  PRIMARY KEY (`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

DROP TABLE IF EXISTS `GUILD_BANK`;
CREATE TABLE `GUILD_BANK` (
  `ID_GUILD` int(10) unsigned NOT NULL,
  `SLOT` tinyint(3) unsigned NOT NULL,
  `ID_BANKBOX` tinyint(3) unsigned NOT NULL,
  `ID_OBJ` mediumint(8) unsigned DEFAULT NULL,
  `AMOUNT` mediumint(8) unsigned DEFAULT NULL,
  PRIMARY KEY (`ID_GUILD`,`SLOT`,`ID_BANKBOX`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


DROP TABLE IF EXISTS `GUILD_INFO`;
CREATE TABLE `GUILD_INFO` (
  `ID_GUILD` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `NAME` varchar(20) NOT NULL,
  `DESCRIPTION` varchar(100) NOT NULL DEFAULT '',
  `ALIGNMENT` tinyint(4) unsigned NOT NULL,
  `CREATION_DATE` datetime NOT NULL,
  `STATUS` tinyint(4) NOT NULL DEFAULT '1',
  `ID_LEADER` int(10) unsigned NOT NULL,
  `ID_RIGHTHAND` int(10) unsigned DEFAULT NULL,
  `MEMBER_COUNT` smallint(6) unsigned DEFAULT NULL,
  `ID_CURRENT_QUEST` int(10) unsigned DEFAULT NULL,
  `CURRENT_QUEST_STAGE` smallint(10) unsigned DEFAULT NULL,
  `QUEST_STARTED_DATE` datetime DEFAULT NULL,
  `CURRENT_QUEST_SECONDS_LEFT` int(10) unsigned NOT NULL DEFAULT '0',
  `CONTRIBUTION_EARNED` int(10) unsigned NOT NULL DEFAULT '0',
  `CONTRIBUTION_AVAILABLE` int(10) unsigned NOT NULL DEFAULT '0',
  `BANK_GOLD` bigint(20) DEFAULT 0,
  `ID_ROLE_NEW_MEMBERS` INT(10) UNSIGNED NULL DEFAULT NULL,
  `RANKING_POINTS` INT(10) UNSIGNED DEFAULT 0,
  PRIMARY KEY (`ID_GUILD`),
  KEY `ID_LEADER_ID_RIGHTHAND` (`ID_LEADER`,`ID_RIGHTHAND`),
  FULLTEXT KEY `NAME` (`NAME`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `GUILD_MEMBER`;
CREATE TABLE `GUILD_MEMBER` (
  `ID_GUILD` int(11) unsigned NOT NULL,
  `ID_USER` int(11) unsigned NOT NULL,
  `JOIN_DATE` datetime DEFAULT NULL,
  `ID_ROLE` int(11) unsigned NOT NULL,
  `ROLE_ASSIGNED_BY` int(11) unsigned NOT NULL,
  `CONTRIBUTION_EARNED` int(11) unsigned NOT NULL DEFAULT '0',
  PRIMARY KEY (`ID_GUILD`,`ID_USER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `GUILD_PERMISSION`;
CREATE TABLE `GUILD_PERMISSION` (
  `ID_PERMISSION` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `KEY` varchar(30) DEFAULT NULL,
  `PERMISSION_NAME` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`ID_PERMISSION`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `GUILD_QUEST_COMPLETED`;
CREATE TABLE `GUILD_QUEST_COMPLETED` (
  `ID_CONTRIBUTION` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `ID_GUILD` int(11) unsigned NOT NULL,
  `ID_QUEST` int(11) unsigned NOT NULL,
  `STARTED_DATE` datetime DEFAULT NULL,
  `COMPLETED_DATE` datetime DEFAULT NULL,
  `MEMBERS_CONTRIBUTED` int(11) unsigned DEFAULT NULL,
  `CONTRIBUTION_GAINED` int(11) unsigned DEFAULT NULL,
  `TOTAL_SECONDS` int(10) unsigned NOT NULL DEFAULT '0',
  KEY `ID_CONTRIBUTION` (`ID_CONTRIBUTION`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `GUILD_ROLE`;
CREATE TABLE `GUILD_ROLE` (
  `ID_ROLE` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `ROLE_NAME` varchar(15) DEFAULT NULL,
  `DELETABLE` bit(1) DEFAULT NULL,
  PRIMARY KEY (`ID_ROLE`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `GUILD_ROLE_ASSIGNED`;
CREATE TABLE `GUILD_ROLE_ASSIGNED` (
  `ID_ROLE` int(11) unsigned NOT NULL,
  `ID_GUILD` int(11) unsigned NOT NULL,
  PRIMARY KEY (`ID_ROLE`,`ID_GUILD`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `GUILD_ROLE_PERMISSION`;
CREATE TABLE `GUILD_ROLE_PERMISSION` (
  `ID_ROLE` int(11) unsigned NOT NULL,
  `ID_PERMISSION` int(11) unsigned NOT NULL,
  PRIMARY KEY (`ID_ROLE`,`ID_PERMISSION`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `GUILD_UPGRADE`;
CREATE TABLE `GUILD_UPGRADE` (
  `ID_GUILD` int(10) unsigned NOT NULL,
  `ID_UPGRADE` int(10) unsigned NOT NULL,
  `UPGRADE_LEVEL` tinyint(3) unsigned DEFAULT '1',
  `UPGRADE_DATE` datetime DEFAULT NULL,
  `UPGRADED_BY` int(10) unsigned DEFAULT NULL,
  `ENABLED` bit(1) DEFAULT b'1',
  PRIMARY KEY (`ID_GUILD`,`ID_UPGRADE`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

DROP TABLE IF EXISTS `user_masteries`;
CREATE TABLE `user_masteries` (
	`ID_USER` INT(10) UNSIGNED NOT NULL,
	`ID_MASTERY` INT(10) UNSIGNED NOT NULL,
	`ID_MASTERY_GROUP` INT(10) UNSIGNED NULL DEFAULT NULL,
	`POINTS_SPENT` INT(10) UNSIGNED NULL DEFAULT NULL,
	`DATE_ADDED` DATETIME NULL DEFAULT NULL,
	PRIMARY KEY (`ID_USER`, `ID_MASTERY`) USING BTREE
)
COLLATE='utf8mb4_general_ci' ENGINE=InnoDB;

DROP TABLE IF EXISTS `GUILD_CURRENT_QUEST_STAGE`;
CREATE TABLE `GUILD_CURRENT_QUEST_STAGE` (
  `ID_QUEST_REQUIREMENT` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `ID_GUILD` int(11) unsigned NOT NULL,
  `REQUIREMENT_TYPE` tinyint(4) UNSIGNED NOT NULL,
  `QUANTITY_COMPLETED` INT(10) UNSIGNED NOT NULL,
  `REQUIREMENT_INDEX` INT(10) UNSIGNED DEFAULT NULL,
  KEY `ID_QUEST_REQUIREMENT` (`ID_QUEST_REQUIREMENT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


DROP TABLE IF EXISTS `QUEST_REQUIREMENT`;
CREATE TABLE `QUEST_REQUIREMENT` (
  `ID_QUEST_REQUIREMENT` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `KEY` varchar(30) DEFAULT NULL,
  KEY `ID_QUEST_REQUIREMENT` (`ID_QUEST_REQUIREMENT`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
