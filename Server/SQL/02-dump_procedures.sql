DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_LoadChar`(in userID int)
MainLabel:BEGIN

    DECLARE IsBanned INT(11);
    DECLARE PunishmentEndDate DATETIME;
    DECLARE PunisherId INT(11);
    DECLARE PunisherName VARCHAR(30);
    DECLARE PunishmentReason VARCHAR(300);
    
    
	-- Check if the character is banned before trying to fetch all the tables
	SELECT 
		UINF.ID_BAN_PUNISHMENT AS PUNISHMENT_ID, 
		IFNULL(USPUNISHER.NAME, '') AS PUNISHED_BY,
		USPUNISHER.ID_USER AS PUNISHER_ID,
		IFNULL(UPUNI.END_DATE, '2050-12-30 00:00:00') AS PUNISHMENT_END_DATE,
		IFNULL(PUNTYP.Description, '') AS PUNISHMENT_REASON
	INTO 
		@IsBanned,
		@PunisherName,
		@PunisherId,
		@PunishmentEndDate,
		@PunishmentReason
	FROM USER_INFO UINF
	LEFT JOIN USER_PUNISHMENT UPUNI
		ON UINF.ID_BAN_PUNISHMENT = UPUNI.ID_PUNISHMENT
	LEFT JOIN PUNISHMENT_TYPE PUNTYP
		ON UPUNI.ID_PUNISHMENT_TYPE = PUNTYP.ID
	LEFT JOIN USER_INFO USPUNISHER
		ON USPUNISHER.ID_USER = UPUNI.ID_PUNISHER		
	WHERE UINF.ID_USER = userID;

	-- If the user is banned, we need to check if the punishment is over.
	IF @IsBanned > 0 THEN
		IF @PunishmentEndDate > NOW() THEN
		
			SELECT 	@IsBanned AS PUNISHMENT_ID,	
						@PunisherId AS PUNISHER_ID, 
						@PunisherName AS PUNISHER_NAME,
						@PunishmentEndDate AS PUNISHMENT_END_DATE,
						@PunishmentReason AS PUNISHMENT_REASON;
						
			-- If the user is still banned then we can skip the retrieval of the rest of the tables
			-- 	as they won't be used by the client. This will give us a small performance saving
			LEAVE MainLabel;			
			
		END IF;
		
		-- we need to unban the character
		CALL sp_UnbanChar(userID, @PunisherId, 1, NOW(), "Automatic UNBAN");
		
	END IF;
	
		-- Return the emtpy dataset as the user is not banned
		SELECT	0 AS PUNISHMENT_ID,	
					0 AS PUNISHER_ID, 
					"" AS PUNISHER_NAME,
					'2050-12-30 00:00:00' AS PUNISHMENT_END_DATE,
					"" AS PUNISHMENT_REASON;
	
		-- User Info (joinable) tables
		-- USER_INFO / USER_STATS / USER_FLAGS 
		-- USER_ATTRIBUTES / USER_FACTION
	SELECT
		UINF.ID_USER,
		UINF.NAME,
		UINF.GENDER,
		UINF.RACE,
		UINF.CLASS,
		UINF.HOME,
		UINF.DESCRIP,
		UINF.PUNISHMENT,
		UINF.HEADING,
		UINF.HEAD,
		UINF.BODY,
		UINF.WEAPON_ANIM,
		UINF.SHIELD_ANIM,
		UINF.HELMET_ANIM,
		UINF.UP_TIME,
		UINF.LAST_POS,
		UINF.WEAPON_SLOT,
		UINF.ARMOUR_SLOT,
		UINF.HELMET_SLOT,
		UINF.SHIELD_SLOT,
		UINF.BOAT_SLOT,
		UINF.MUNITION_SLOT,
		UINF.SACKPACK_SLOT,
		UINF.RING_SLOT,
		UINF.TRAINNING_TIME,
		USTAT.ORO,
		USTAT.ORO_BANCO,
		USTAT.HP_MAX,
		USTAT.HP_MIN,
		USTAT.STAMINA_MAX,
		USTAT.STAMINA_MIN,
		USTAT.MANA_MAX,
		USTAT.MANA_MIN,
		USTAT.HIT_MAX,
		USTAT.HIT_MIN,
		USTAT.AGUA_MAX,
		USTAT.AGUA_MIN,
		USTAT.HAMBRE_MAX,
		USTAT.HAMBRE_MIN,
		USTAT.SKILLS,
		USTAT.EXP,
		USTAT.NIVEL,
		USTAT.EXP_NEXT,
		USTAT.USERS_KILLED,
		USTAT.NPCS_KILLED,
		USTAT.MASTERY_POINTS,
		USTAT.RANKING_POINTS,
		USTAT.DUELOS_GANADOS,
		USTAT.DUELOS_PERDIDOS,
		USTAT.ORO_DUELOS,
		UINF.GUILD_ID,
		UINF.REQUESTING_GUILD,
		IFNULL(UINF.GUILD_REJECT_DETAIL, '') AS GUILD_REJECT_DETAIL,
		UFLAG.MUERTO,
		UFLAG.ESCONDIDO,
		UFLAG.HAMBRE,
		UFLAG.SED,
		UFLAG.DESNUDO,
		UFLAG.BAN,
		UFLAG.NAVEGANDO,
		UFLAG.ENVENENADO,
		UFLAG.PARALIZADO,
		UFLAG.LAST_MAP,
		UFLAG.LAST_TAMED_PET,
		UATTR.STRENGHT,
		UATTR.DEXERITY,
		UATTR.INTELLIGENCE,
		UATTR.CHARISM,
		UATTR.HEALTH,
		UFACT.ALIGNMENT,
		UFACT.ARMY,
		UFACT.CHAOS,
		UFACT.NEUTRAL_KILLED,
		UFACT.CITY_KILLED,
		UFACT.CRI_KILLED,
		UFACT.CHAOS_ARMOUR_GIVEN,
		UFACT.ARMY_ARMOUR_GIVEN,
		UFACT.CHAOS_EXP_GIVEN,
		UFACT.ARMY_EXP_GIVEN,
		UFACT.CHAOS_REWARD_GIVEN,
		UFACT.ARMY_REWARD_GIVEN,
		UFACT.NUM_SIGNS,
		UFACT.SIGNING_LEVEL,
		UFACT.SIGNING_DATE,
		UFACT.SIGNING_KILLED,
		UFACT.NEXT_REWARD,
		UFACT.ROYAL_COUNCIL,
		UFACT.CHAOS_COUNCIL
--		UINF.ID_BAN_PUNISHMENT,
--		IFNULL(USPUNISHER.NAME, '') AS PUNISHED_BY,
--		IFNULL(PUNTYP.Description, '') AS PUNISHMENT_REASON,
--		IFNULL(UPUNI.END_DATE, '2050-12-30 00:00:00') AS PUNISHMENT_END_DATE
	FROM
		USER_INFO UINF
			INNER JOIN USER_STATS USTAT
				ON UINF.ID_USER = USTAT.ID_USER
			INNER JOIN USER_FLAGS UFLAG
				ON UINF.ID_USER = UFLAG.ID_USER
			INNER JOIN USER_ATTRIBUTES UATTR
				ON UINF.ID_USER = UATTR.ID_USER
			INNER JOIN USER_FACTION UFACT
				ON UINF.ID_USER = UFACT.ID_USER
--			LEFT JOIN USER_PUNISHMENT UPUNI
--				ON UINF.ID_BAN_PUNISHMENT = UPUNI.ID_PUNISHMENT
--			LEFT JOIN PUNISHMENT_TYPE PUNTYP
--				ON UPUNI.ID_PUNISHMENT_TYPE = PUNTYP.ID
--			LEFT JOIN USER_INFO USPUNISHER
--				ON USPUNISHER.ID_USER = UPUNI.ID_PUNISHER
	WHERE
		UINF.ID_USER = userID;
		
	-- User Inventory
	SELECT
		SLOT,
		OBJ_INDEX,
		AMOUNT,
		EQUIPPED
	FROM
		USER_INVENTORY
	WHERE ID_USER = userID;

	-- User Bank
	SELECT
		SLOT, OBJ_INDEX, AMOUNT
	FROM
		USER_BANK
	WHERE
		ID_USER = userID;

	-- User Spells
	SELECT
		SLOT, SPELL_INDEX
	FROM
		USER_SPELLS
	WHERE
		ID_USER = userID;

	-- User Skills
	SELECT
		SKILL,
		NATURAL_AMOUNT,
		ASSIGNED_AMOUNT,
		SKILL_EXP_NEXT_LEVEL,
		SKILL_EXP
	FROM
		USER_SKILLS
	WHERE
		ID_USER = userID;

	-- User Pets
	SELECT
		NUM_PET, NPC_INDEX, NPC_TYPE, NPC_LIFE
	FROM
		USER_PETS
	WHERE
		ID_USER = userID;

	-- User Masteries
	SELECT 	ID_USER,
				ID_MASTERY_GROUP,
				GROUP_CONCAT(ID_MASTERY SEPARATOR ',') AS MASTERIES
	FROM USER_MASTERIES
	WHERE ID_USER = UserID
	GROUP BY ID_MASTERY_GROUP;
		
	END$$

	DELIMITER ;


	DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_DeleteChar` (in userID int)
	BEGIN
		-- IglorioN: Stored procedure to delete all character info in database
		
		DELETE FROM account_char_share
	WHERE 
		ID_USER = userID;
		
		DELETE FROM guild_member 
	WHERE
		ID_USER = userID;
		
		DELETE FROM statictics 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_attributes 
	WHERE
		ID_USER = userID;
		
		
		DELETE FROM user_bank 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_ban_detail 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_faction 
	WHERE
		ID_USER = userID;
		
		DELETE FROM USER_FLAGS 
	WHERE
		ID_USER = userID;  
		
		DELETE FROM user_info 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_inventory 
	WHERE
		ID_USER = userID;
		
		
		DELETE FROM user_messages 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_pets 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_polls 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_punishment 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_skills 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_spells 
	WHERE
		ID_USER = userID;
		
		DELETE FROM user_stats 
	WHERE
		ID_USER = userID;      
		
END$$

DELIMITER ;


DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_SetCharacterLoggedIn`(IN userID int, IN userIP varchar(15), IN isConnectionEvent boolean, IN currentConnectionEventId int)
	BEGIN

		-- Set's the user as connected or disconnected

		IF isConnectionEvent = true THEN
			UPDATE USER_INFO
			SET LOGGED = 1
			WHERE ID_USER = userID;

			INSERT INTO user_connections (CONNECTION_DATE, ID_USER, IP)
			VALUES (NOW(), userID, userIP);

			SELECT LAST_INSERT_ID() as ID_EVENT;
		ELSE

			UPDATE user_info
			SET LOGGED = 0
			WHERE ID_USER = userID;

			if (currentConnectionEventId != 0) THEN
				UPDATE user_connections
				SET DISCONNECTION_DATE = NOW()
				WHERE ID_EVENT = currentConnectionEventId;
			END IF;

			SELECT currentConnectionEventId as ID_EVENT;

		END IF;


	END $$
DELIMITER ;

DELIMITER $$
CREATE OR REPLACE procedure `sp_SaveCharacterHeader`(IN UserID int, IN UserName varchar(30), IN Gender tinyint,
                                                        IN Race tinyint, IN Class tinyint, IN Home tinyint,
                                                        IN CharDescription varchar(50), IN Punishment tinyint,
                                                        IN Heading tinyint, IN Head mediumint, IN Body mediumint,
                                                        IN WeaponAnim mediumint, IN ShieldAnim mediumint,
                                                        IN HelmetAnim mediumint, IN Uptime int, IN LastIP varchar(20),
                                                        IN LastPoss varchar(11), IN WeaponSlot tinyint,
                                                        IN ArmorSlot tinyint, IN HelmetSlot tinyint,
                                                        IN ShieldSlot tinyint, IN BoatSlot tinyint, IN AmmoSlot tinyint,
                                                        IN BackpackSlot tinyint, IN RingSlot tinyint,
                                                        IN GuildId mediumint, IN RequestingGuild mediumint,
                                                        IN IsBanned tinyint, IN LastPunishmentId int,
                                                        IN TrainingTime int, IN AccountId int, IN IsDead tinyint,
                                                        IN IsHidding tinyint, IN Hunger tinyint, IN Thirst tinyint,
                                                        IN Naked tinyint, IN Sailing tinyint, IN Poisoned tinyint,
                                                        IN Paralized tinyint, IN LastMap smallint,
                                                        IN LastTammedPet mediumint, IN Gold int, IN GoldBank int,
                                                        IN HpMax smallint, IN HpMin smallint, IN StaminaMax smallint,
                                                        IN StaminaMin smallint, IN ManaMax smallint, IN ManaMin smallint, 
                                                        IN ThirstMax smallint, IN ThirstMin smallint,
                                                        IN HungerMax smallint, IN HungerMin smallint,
                                                        IN Skills smallint, IN Exp bigint, IN Level tinyint,
                                                        IN ExpNextLevel bigint, IN UsersKilled mediumint,
                                                        IN NpcsKilled mediumint, IN RankingPoints int,
                                                        IN MasteryPoints int, IN DuelsWon mediumint,
                                                        IN DuelsLost mediumint, IN DuelsGoldWon int, 
														IN Alignment tinyint, IN IsArmy tinyint,
                                                        IN IsChaos tinyint, IN NeutralsKilled mediumint, IN CitizensKilled mediumint,
                                                        IN CriminalsKilled mediumint, IN ChaosArmorGiven tinyint,
                                                        IN ArmyArmorGiven tinyint, IN ChaosExpGiven tinyint,
                                                        IN ArmyExpGiven tinyint, IN ChaosRewardGiven tinyint,
                                                        IN ArmyRewardGiven tinyint, IN NumSigns int,
                                                        IN SigningLevel tinyint, IN SigningDate datetime,
                                                        IN SigningKilled mediumint, IN NextReward int,
                                                        IN IsRoyalCouncil tinyint, IN IsChaosCouncil tinyint)
MainLabel:BEGIN


DECLARE NewCharacterId INT;
    
	IF (UserID = 0) THEN
		-- Creation of a character
		INSERT INTO user_info (
			`NAME`, GENDER, RACE, CLASS, HOME, DESCRIP, PUNISHMENT, HEADING, HEAD, BODY,WEAPON_ANIM,
			SHIELD_ANIM, HELMET_ANIM, UP_TIME, LAST_IP, LAST_POS, WEAPON_SLOT, ARMOUR_SLOT, HELMET_SLOT,
			SHIELD_SLOT, BOAT_SLOT, MUNITION_SLOT, SACKPACK_SLOT, RING_SLOT, GUILD_ID, REQUESTING_GUILD,
			BANNED, ID_BAN_PUNISHMENT, TRAINNING_TIME, ID_ACCOUNT, CREATION_DATE
			)
		VALUES (
				UserName, Gender, Race, Class, Home, CharDescription, Punishment, heading, Head, Body, WeaponAnim,
				ShieldAnim, HelmetAnim, Uptime, LastIP, LastPoss, WeaponSlot, ArmorSlot, HelmetSlot, 
				ShieldSlot, BoatSlot, AmmoSlot, BackpackSlot, RingSlot, GuildId, RequestingGuild, 
				IsBanned, LastPunishmentId, TrainingTime, AccountId, NOW()
				);
    
		-- Get the new character ID based on the last inserted row in USER_INDEX
		SET NewCharacterId := LAST_INSERT_ID();
        
        INSERT INTO USER_FLAGS (`ID_USER`,`MUERTO`,`ESCONDIDO`,`HAMBRE`,`SED`,`DESNUDO`,`NAVEGANDO`,`ENVENENADO`,
								`PARALIZADO`,`LAST_MAP`,`LAST_TAMED_PET`)
						VALUES (NewCharacterId,
								IsDead,
                                IsHidding,
                                Hunger,
                                Thirst,
                                Naked,
                                Sailing,
                                Poisoned,
                                Paralized,
                                LastMap,
                                LastTammedPet);
                                
                                
		INSERT INTO USER_STATS (`ID_USER`,`ORO`,`ORO_BANCO`,`HP_MAX`,`HP_MIN`,`STAMINA_MAX`,`STAMINA_MIN`,`MANA_MAX`,
								`MANA_MIN`,`AGUA_MAX`,`AGUA_MIN`,`HAMBRE_MAX`,`HAMBRE_MIN`,
								`SKILLS`,`EXP`,`NIVEL`,`EXP_NEXT`,`USERS_KILLED`,`NPCS_KILLED`,`RANKING_POINTS`,
								`MASTERY_POINTS`,`DUELOS_GANADOS`,`DUELOS_PERDIDOS`,`ORO_DUELOS`)
					VALUES		(NewCharacterId,
								Gold,
                                GoldBank,
                                HpMax,
                                HpMin,
                                StaminaMax,
                                StaminaMin,
                                ManaMax,
                                ManaMin,
                                ThirstMax,
                                ThirstMin,
                                HungerMax,
                                HungerMin,
                                Skills,
                                Exp,
                                Level,
                                ExpNextLevel,
                                UsersKilled,
                                NpcsKilled,
                                RankingPoints,
                                MasteryPoints,
                                DuelsWon,
                                DuelsLost,
                                DuelsGoldWon);
                                
		INSERT INTO user_faction (`ID_USER`,`ALIGNMENT`,`ARMY`,`CHAOS`,`NEUTRAL_KILLED`,`CITY_KILLED`,`CRI_KILLED`,`CHAOS_ARMOUR_GIVEN`,`ARMY_ARMOUR_GIVEN`,
									`CHAOS_EXP_GIVEN`,`ARMY_EXP_GIVEN`,`CHAOS_REWARD_GIVEN`,`ARMY_REWARD_GIVEN`,
									`NUM_SIGNS`,`SIGNING_LEVEL`,`SIGNING_DATE`,`SIGNING_KILLED`,`NEXT_REWARD`,
									`ROYAL_COUNCIL`,`CHAOS_COUNCIL`)
						VALUES		(NewCharacterId,
												Alignment,
												IsArmy,
                                    IsChaos,
									NeutralsKilled,
                                    CitizensKilled,
                                    CriminalsKilled,
                                    ChaosArmorGiven,
                                    ArmyArmorGiven,
                                    ChaosExpGiven,
                                    ArmyExpGiven,
                                    ChaosRewardGiven,
                                    ArmyRewardGiven,
                                    NumSigns,
                                    SigningLevel,
                                    SigningDate,
                                    SigningKilled,
                                    NextReward,
                                    IsRoyalCouncil,
                                    IsChaosCouncil);

	ELSE
		-- Update of an existing character
		-- USER_INFO
		UPDATE `user_info`
		SET
			`NAME` = UserName,
			`GENDER` = Gender,
			`RACE` = Race,
			`CLASS` = Class,
			`HOME` = Home,
			`DESCRIP` = CharDescription,
			`PUNISHMENT` = Punishment,
			`HEADING` = Heading,
			`HEAD` = Head,
			`BODY` = Body,
			`WEAPON_ANIM` = WeaponAnim,
			`SHIELD_ANIM` = ShieldAnim,
			`HELMET_ANIM` = HelmetAnim,
			`UP_TIME` = Uptime,
			`LAST_IP` = LastIP,
			`LAST_POS` = LastPoss,
			`WEAPON_SLOT` = WeaponSlot,
			`ARMOUR_SLOT` = ArmorSlot,
			`HELMET_SLOT` = HelmetSlot,
			`SHIELD_SLOT` = ShieldSlot,
			`BOAT_SLOT` = BoatSlot,
			`MUNITION_SLOT` = AmmoSlot,
			`SACKPACK_SLOT` = BackpackSlot,
			`RING_SLOT` = RingSlot,
			`GUILD_ID` = GuildId,
			`REQUESTING_GUILD` = RequestingGuild,
			`BANNED` = IsBanned,
			`ID_BAN_PUNISHMENT` = LastPunishmentId,
			`TRAINNING_TIME` = TrainingTime,
			`ID_ACCOUNT` = AccountId
		WHERE `ID_USER` = UserID;
        
        UPDATE `user_flags`
		SET
			`MUERTO` = IsDead,
			`ESCONDIDO` = IsHidding,
			`HAMBRE` = Hunger,
			`SED` = Thirst,
			`DESNUDO` = Naked,
			`NAVEGANDO` = Sailing,
			`ENVENENADO` = Poisoned,
			`PARALIZADO` = Paralized,
			`LAST_MAP` = LastMap,
			`LAST_TAMED_PET` = LastTammedPet
		WHERE `ID_USER` = UserID;
        
        UPDATE `user_stats`
		SET
			`ORO` = Gold,
			`ORO_BANCO` = GoldBank,
			`HP_MAX` = HpMax,
			`HP_MIN` = HpMin,
			`STAMINA_MAX` = StaminaMax,
			`STAMINA_MIN` = StaminaMin,
			`MANA_MAX` = ManaMax,
			`MANA_MIN` = ManaMin,
			`AGUA_MAX` = ThirstMax,
			`AGUA_MIN` = ThirstMin,
			`HAMBRE_MAX` = HungerMax,
			`HAMBRE_MIN` = HungerMin,
			`SKILLS` = Skills,
			`EXP` = Exp,
			`NIVEL` = Level,
			`EXP_NEXT` = ExpNextLevel,
			`USERS_KILLED` = UsersKilled,
			`NPCS_KILLED` = NpcsKilled,
			`RANKING_POINTS` = RankingPoints,
			`MASTERY_POINTS` = MasteryPoints,
			`DUELOS_GANADOS` = DuelsWon,
			`DUELOS_PERDIDOS` = DuelsLost,
			`ORO_DUELOS` = DuelsGoldWon
		WHERE `ID_USER` = UserID;

		UPDATE `user_faction`
		SET
			`ALIGNMENT` = Alignment,
			`ARMY` = IsArmy,
			`CHAOS` = IsChaos,
			`NEUTRAL_KILLED` = NeutralsKilled,
			`CITY_KILLED` = CitizensKilled,
			`CRI_KILLED` = CriminalsKilled,
			`CHAOS_ARMOUR_GIVEN` = ChaosArmorGiven,
			`ARMY_ARMOUR_GIVEN` = ArmyArmorGiven,
			`CHAOS_EXP_GIVEN` = ChaosExpGiven,
			`ARMY_EXP_GIVEN` = ArmyExpGiven,
			`CHAOS_REWARD_GIVEN` = ChaosRewardGiven,
			`ARMY_REWARD_GIVEN` = ArmyRewardGiven,
			`NUM_SIGNS` = NumSigns,
			`SIGNING_LEVEL` = SigningLevel,
			`SIGNING_DATE` = SigningDate,
			`SIGNING_KILLED` = SigningKilled,
			`NEXT_REWARD` = NextReward,
			`ROYAL_COUNCIL` = IsRoyalCouncil,
			`CHAOS_COUNCIL` = IsChaosCouncil
		WHERE `ID_USER` = UserID;
        
		SET NewCharacterId := UserID;
	END IF;
    
    
    SELECT NewCharacterId as ID_USER;



END$$
DELIMITER ;

DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_LoadGuild`(IN GuildID int)
BEGIN

/* Guild Info */
SELECT 
	GUILD_INFO.ID_GUILD,
	GUILD_INFO.`NAME`,
	GUILD_INFO.`DESCRIPTION`,
	GUILD_INFO.ALIGNMENT,
	GUILD_INFO.CREATION_DATE,
	GUILD_INFO.`STATUS`,
	GUILD_INFO.ID_LEADER,
	GUILD_INFO.ID_RIGHTHAND,
	GUILD_INFO.MEMBER_COUNT,
	GUILD_INFO.ID_CURRENT_QUEST,
	GUILD_INFO.QUEST_STARTED_DATE,
	GUILD_INFO.CONTRIBUTION_EARNED,
	GUILD_INFO.CONTRIBUTION_AVAILABLE,
	GUILD_INFO.BANK_GOLD,
	GUILD_INFO.ID_ROLE_NEW_MEMBERS,
	GUILD_INFO.CURRENT_QUEST_STAGE,
	GUILD_INFO.CURRENT_QUEST_SECONDS_LEFT,
	RANKING_POINTS
FROM GUILD_INFO
WHERE ID_GUILD = GuildID;

SELECT 
	GUILD_ROLE_ASSIGNED.ID_ROLE,
    GUILD_ROLE.ROLE_NAME,
    GUILD_ROLE.DELETABLE,
	GUILD_ROLE_ASSIGNED.ID_GUILD
FROM GUILD_ROLE_ASSIGNED
JOIN GUILD_ROLE ON GUILD_ROLE_ASSIGNED.ID_ROLE = GUILD_ROLE.ID_ROLE
WHERE GUILD_ROLE_ASSIGNED.ID_GUILD = GuildID;

SELECT 
	GUILD_MEMBER.ID_GUILD,
    GUILD_MEMBER.ID_USER,
    USER_INFO.`NAME`,
    GUILD_MEMBER.JOIN_DATE,
    GUILD_MEMBER.ID_ROLE,
    GUILD_MEMBER.ROLE_ASSIGNED_BY,
    GUILD_MEMBER.CONTRIBUTION_EARNED
FROM GUILD_MEMBER
JOIN USER_INFO ON GUILD_MEMBER.ID_USER = USER_INFO.ID_USER
WHERE ID_GUILD = GuildID;

SELECT
	GUILD_ROLE_PERMISSION.ID_ROLE,
    GUILD_ROLE_PERMISSION.ID_PERMISSION,
    GUILD_PERMISSION.KEY
FROM GUILD_ROLE_PERMISSION
JOIN GUILD_ROLE_ASSIGNED ON GUILD_ROLE_PERMISSION.ID_ROLE = GUILD_ROLE_ASSIGNED.ID_ROLE
JOIN GUILD_PERMISSION ON GUILD_PERMISSION.ID_PERMISSION = GUILD_ROLE_PERMISSION.ID_PERMISSION
WHERE GUILD_ROLE_ASSIGNED.ID_GUILD = GuildID;

SELECT
	GUILD_QUEST_COMPLETED.ID_CONTRIBUTION,
    GUILD_QUEST_COMPLETED.ID_GUILD,
    GUILD_QUEST_COMPLETED.ID_QUEST,
    GUILD_QUEST_COMPLETED.COMPLETED_DATE,
    GUILD_QUEST_COMPLETED.MEMBERS_CONTRIBUTED,
    GUILD_QUEST_COMPLETED.CONTRIBUTION_GAINED
FROM GUILD_QUEST_COMPLETED
WHERE ID_GUILD = GuildID;

SELECT
	GUILD_UPGRADE.ID_GUILD,
	GUILD_UPGRADE.ID_UPGRADE,
	GUILD_UPGRADE.UPGRADE_LEVEL,
	GUILD_UPGRADE.UPGRADE_DATE,
	GUILD_UPGRADE.UPGRADED_BY,
	GUILD_UPGRADE.ENABLED
FROM GUILD_UPGRADE
WHERE GUILD_UPGRADE.ID_GUILD= GuildID;

SELECT
	GUILD_BANK.ID_GUILD,
    GUILD_BANK.SLOT,
	GUILD_BANK.ID_BANKBOX,
    GUILD_BANK.ID_OBJ,
    GUILD_BANK.AMOUNT
FROM GUILD_BANK
WHERE GUILD_BANK.ID_GUILD = GuildID;
SELECT
	GUILD_CURRENT_QUEST_STAGE.REQUIREMENT_TYPE,
	GUILD_CURRENT_QUEST_STAGE.QUANTITY_COMPLETED,
	GUILD_CURRENT_QUEST_STAGE.REQUIREMENT_INDEX
FROM GUILD_CURRENT_QUEST_STAGE
WHERE ID_GUILD = GuildID;


END$$
DELIMITER ;


DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_CreateGuild`(IN GuildName varchar(30), IN LeaderID int, IN Alignment tinyint, IN RankingPoints int)
BEGIN

DECLARE NewGuildId, LeaderRoleId, MemberRoleId INT;

SET LeaderRoleId = 1;

-- Let's insert the Member's role first so we can use it after
INSERT INTO GUILD_ROLE
(	ROLE_NAME,	DELETABLE )
VALUES
(	"Recluta", 0 );

SET MemberRoleId := LAST_INSERT_ID();

-- Insert the GUILD basic data
INSERT INTO GUILD_INFO
(
	`NAME`,
	ALIGNMENT,
	CREATION_DATE,
	`STATUS`,
	ID_LEADER,
	ID_RIGHTHAND,
	MEMBER_COUNT,
	ID_ROLE_NEW_MEMBERS,
	RANKING_POINTS
)
VALUES
(
	GuildName,
	Alignment,
	NOW(),			/* Creation Date */
	1, 				/* Status */
	LeaderID,
	NULL, 			/* id right hand */
	1 				/* Starting member count */,
	MemberRoleId,
	RankingPoints
);

SET NewGuildId := LAST_INSERT_ID();

INSERT INTO GUILD_MEMBER
(
	ID_GUILD,
	ID_USER,
	JOIN_DATE,
	ID_ROLE,
	ROLE_ASSIGNED_BY
)
VALUES
(	NewGuildId, LeaderId, NOW(), LeaderRoleId,	LeaderId );
 
UPDATE USER_INFO SET GUILD_ID = NewGuildId WHERE ID_USER = LeaderId;

INSERT INTO `GUILD_ROLE_ASSIGNED`
(
	ID_ROLE,
	ID_GUILD
)
VALUES
(1, NewGuildId),
(2, NewGuildId),
(MemberRoleId, NewGuildId);

CALL sp_LoadGuild(NewGuildId);

END$$
DELIMITER ;

DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_GuildMemberAdd`(IN GuildID int, IN MemberRequestID int, in TargetRequestID int, IN RolMemberID int)
BEGIN
DECLARE QtyMember INT;
	INSERT INTO GUILD_MEMBER (ID_GUILD, ID_USER, JOIN_DATE, ID_ROLE, ROLE_ASSIGNED_BY)
	VALUES	(	GuildID, TargetRequestID, NOW(), RolMemberID,	 MemberRequestID);
	UPDATE USER_INFO SET GUILD_ID = GuildID WHERE ID_USER = TargetRequestID;
    
	SELECT count(ID_GUILD) into QtyMember FROM GUILD_MEMBER WHERE ID_GUILD=GuildID;
	UPDATE GUILD_INFO SET MEMBER_COUNT = QtyMember WHERE GUILD_INFO.ID_GUILD= GuildID;
	
	SELECT GUILD_MEMBER.ID_GUILD, GUILD_MEMBER.ID_USER, USER_INFO.`NAME`, GUILD_MEMBER.JOIN_DATE, GUILD_MEMBER.ID_ROLE, GUILD_MEMBER.ROLE_ASSIGNED_BY, GUILD_MEMBER.CONTRIBUTION_EARNED
	FROM GUILD_MEMBER JOIN USER_INFO ON GUILD_MEMBER.ID_USER = USER_INFO.ID_USER WHERE GUILD_MEMBER.ID_GUILD=GuildID AND GUILD_MEMBER.ID_USER=TargetRequestID;
END$$
DELIMITER ;

DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_CreateRole`(IN GuildId int, IN RoleName varchar(15), IN Permissions varchar(255))
BEGIN

DECLARE NewRoleId INT;

INSERT INTO GUILD_ROLE(ROLE_NAME, DELETABLE) VALUES (RoleName, 1);

SET NewRoleId := LAST_INSERT_ID();

INSERT INTO GUILD_ROLE_ASSIGNED(ID_ROLE, ID_GUILD)
 VALUES(NewRoleId, GuildId);

INSERT INTO GUILD_ROLE_PERMISSION
SELECT NewRoleId, ID_PERMISSION
FROM GUILD_PERMISSION
WHERE FIND_IN_SET(`KEY`,Permissions) > 0;

SELECT NewRoleId AS ID_ROLE;
END$$
DELIMITER ;


DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_ModifyRole`(IN RoleId int, IN RoleName varchar(30), IN Permissions varchar(255))
BEGIN

UPDATE GUILD_ROLE
SET ROLE_NAME = RoleName
WHERE ID_ROLE = RoleId;

DELETE FROM GUILD_ROLE_PERMISSION WHERE ID_ROLE = RoleId;

INSERT INTO GUILD_ROLE_PERMISSION
SELECT RoleId, ID_PERMISSION
FROM GUILD_PERMISSION
WHERE FIND_IN_SET(`KEY`,Permissions) > 0;

END$$
DELIMITER ;

DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_GuildMemberUpdate`(IN GuildID int, IN MemberRequestID int, in TargetRequestID int, IN RoleMemberID int, IN Contribution int)
BEGIN
	UPDATE GUILD_MEMBER 
	SET ID_ROLE=RoleMemberID, ROLE_ASSIGNED_BY = MemberRequestID, CONTRIBUTION_EARNED=Contribution
	WHERE ID_GUILD=GuildID AND ID_USER=TargetRequestID;
END$$
DELIMITER ;

DELIMITER $$
CREATE OR REPLACE PROCEDURE `sp_GuildMemberDelete`(IN GuildID int, in TargetRequestID int)
BEGIN
	DECLARE QtyMember INT;
	DELETE FROM GUILD_MEMBER WHERE ID_GUILD=GuildID AND ID_USER=TargetRequestID;			
	UPDATE USER_INFO SET GUILD_ID = 0 WHERE ID_USER = TargetRequestID;

	SELECT count(ID_GUILD) into QtyMember FROM GUILD_MEMBER WHERE ID_GUILD=GuildID;
	UPDATE GUILD_INFO SET MEMBER_COUNT = QtyMember WHERE GUILD_INFO.ID_GUILD= GuildID;
END$$
DELIMITER ;
DELIMITER ;

DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_AddUserMastery`(IN userID int, IN masteryID int, IN groupID INT, IN pointsSpent INT)
	BEGIN

		INSERT INTO USER_MASTERIES(ID_USER, ID_MASTERY, ID_MASTERY_GROUP, POINTS_SPENT, DATE_ADDED) VALUES (userID, masteryID, groupID, pointsSpent, NOW());
		
	END $$
DELIMITER ;

DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_AcceptGuildQuest`(IN GuildId int, IN QuestId int, IN StageId int, IN StartDate DATETIME, IN SecondsLeft int)
	BEGIN

		DELETE FROM GUILD_CURRENT_QUEST_STAGE WHERE ID_GUILD = GuildId;

		UPDATE GUILD_INFO   
		SET QUEST_STARTED_DATE = StartDate,
			ID_CURRENT_QUEST = QuestId,
			CURRENT_QUEST_STAGE = StageId,
			CURRENT_QUEST_SECONDS_LEFT = SecondsLeft
		WHERE ID_GUILD = GuildId;

END$$
DELIMITER ;

DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_DeleteGuildCurrentQuest`(IN GuildId int)
	BEGIN

		DELETE FROM GUILD_CURRENT_QUEST_STAGE WHERE ID_GUILD = GuildId;

		UPDATE GUILD_INFO 
  		SET QUEST_STARTED_DATE = NULL, 
		  	ID_CURRENT_QUEST = NULL,
			CURRENT_QUEST_STAGE = NULL
            WHERE ID_GUILD = GuildId;

END$$
DELIMITER ;

DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_FinishGuildCurrentQuest`(IN GuildId int, IN QuestId int, IN Members INT, IN Contribution int,IN TotalSeconds int)
	BEGIN
	
	DECLARE StartedDate DATETIME;	
	
	SELECT StartedDate = QUEST_STARTED_DATE
		FROM GUILD_INFO
		WHERE ID_GUILD = GuildId;
	
	INSERT INTO GUILD_QUEST_COMPLETED (
		ID_GUILD, 
		ID_QUEST, 
		STARTED_DATE,
		COMPLETED_DATE, 
		MEMBERS_CONTRIBUTED, 
		CONTRIBUTION_GAINED,
		TOTAL_SECONDS
		) 
	VALUES ( 
		GuildId,
		QuestId,
		StartedDate,
		NOW(),
		Members,
		Contribution,
		TotalSeconds
			);

	CALL sp_DeleteGuildCurrentQuest(GuildId);
END$$
DELIMITER ;

DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_UnbanChar`(
	IN `userID` INT,
	IN `punisherID` INT,
	IN `punishmentTypeID` INT,
	IN `unbanDate` DATETIME,
	IN `reason` VARCHAR(255))
	BEGIN
		
		-- Update the user's table
		UPDATE USER_INFO
		SET	BANNED = 0,
				ID_BAN_PUNISHMENT = 0
		WHERE ID_USER = userID;
		
		-- Add the record to the list of user punishments
		INSERT INTO USER_PUNISHMENT (	ID_USER,
												ID_PUNISHER,
												ID_PUNISHMENT_TYPE,
												EVENT_DATE,
												END_DATE,
												REASON,
												ADMIN_NOTES)

		VALUES (	userID,
				punisherID,
				punishmentTypeID,
				unbanDate,
				unbanDate,
				reason,
				'');
	END $$
DELIMITER ;



DELIMITER $$

	CREATE OR REPLACE PROCEDURE `sp_AddCharacterPunishment`(
		IN `UserID` INT,
		IN `PunisherID` INT,
		IN `PunishmentTypeID` INT,
		IN `Reason` VARCHAR(255),
		IN `AdminNotes` VARCHAR(255))
	BEGIN

		DECLARE SamePunishmentCount INT;
		DECLARE PunishmentSeverity INT;
		DECLARE PunishmentBaseType INT;
		DECLARE PunishmentEndDate DATETIME;	
		DECLARE ForcedPunishmentTypeID INT;
		DECLARE ForcedPunishmentSeverity INT;
		DECLARE ForcedPunishmentSubtype INT;
		DECLARE ForcedPunishmentBaseType INT;
		DECLARE AddBan INT;
		
		DECLARE CurrentPunishmentCount INT;
		
		SET @CurrentPunishmentCount = (SELECT COUNT(1) AS CNT
											FROM USER_PUNISHMENT UP
											WHERE ID_PUNISHMENT_TYPE = PunishmentTypeID 
												AND ID_USER = UserID);
		
		
		SELECT 	T.BASE_TYPE,
					TR.PUNISHMENT_SEVERITY,
					TR.ADD_BAN,
					TR.NEXT_PUNISHMENT_ID,
					NEXT_T.BASE_TYPE
		INTO 		@PunishmentBaseType,
					@PunishmentSeverity,
					@AddBan,
					@ForcedPunishmentTypeID,
					@ForcedPunishmentBaseType
		FROM PUNISHMENT_TYPE_RULES TR  
		INNER JOIN PUNISHMENT_TYPE T 
			ON TR.ID_PUNISHMENT_TYPE = T.ID  
		LEFT JOIN PUNISHMENT_TYPE NEXT_T
			ON TR.NEXT_PUNISHMENT_ID = NEXT_T.ID
		Where TR.ID_PUNISHMENT_TYPE = PunishmentTypeID 
			-- AND (T.BASE_TYPE <> 4 AND T.ENABLED = 1 or T.BASE_TYPE = 4)
			AND PUNISHMENT_COUNT > @CurrentPunishmentCount
		ORDER BY PUNISHMENT_COUNT 
		LIMIT 1;
		
		
		-- 1 is JAIL that is increased in minutes. Others are increased in days.
		IF @PunishmentBaseType = 1 THEN
			SET @PunishmentEndDate = (DATE_ADD(NOW(), INTERVAL @PunishmentSeverity MINUTE));
		ELSEIF @PunishmentBaseType = 2 THEN
			SET @PunishmentEndDate = (DATE_ADD(NOW(), INTERVAL @PunishmentSeverity DAY));
		ELSE
			SET @PunishmentEndDate = NOW();
		END IF;

		-- Add the record to the list of user punishments
		INSERT INTO USER_PUNISHMENT (	ID_USER,
												ID_PUNISHER,
												ID_PUNISHMENT_TYPE,
												EVENT_DATE,
												END_DATE,
												REASON,
												ADMIN_NOTES)
		VALUES (	UserID,
					PunisherID,
					PunishmentTypeID,
					NOW(),
					@PunishmentEndDate,
					Reason,
					AdminNotes);

		-- If the current punishment accumulation leads to a BAN, then we need to apply
		-- add another punishment. Ie: Multiple consecutive jails of the same type
		-- leads to a ban for some time.
		IF @AddBan = 1 THEN
										
				
				SET @CurrentPunishmentCount = (SELECT COUNT(1) AS CNT
													FROM USER_PUNISHMENT UP
													WHERE ID_PUNISHMENT_TYPE = @ForcedPunishmentTypeID 
														AND ID_USER = UserID);
				
				SELECT 	T.BASE_TYPE,
					TR.PUNISHMENT_SEVERITY
				INTO 		@ForcedPunishmentBaseType,
							@ForcedPunishmentSeverity
				FROM PUNISHMENT_TYPE_RULES TR  
				INNER JOIN PUNISHMENT_TYPE T 
					ON TR.ID_PUNISHMENT_TYPE = T.ID  
				Where TR.ID_PUNISHMENT_TYPE = @ForcedPunishmentTypeID 
					AND PUNISHMENT_COUNT >= @CurrentPunishmentCount
				ORDER BY PUNISHMENT_COUNT 
				LIMIT 1;
				
			-- This is the punishment of the 
			IF @ForcedPunishmentBaseType = 1 THEN
				SET @PunishmentEndDate = (DATE_ADD(NOW(), INTERVAL @ForcedPunishmentSeverity MINUTE));
			ELSEIF @ForcedPunishmentBaseType = 2 THEN
				SET @PunishmentEndDate = (DATE_ADD(NOW(), INTERVAL @ForcedPunishmentSeverity DAY));
			ELSE
				SET @PunishmentEndDate = NOW();
			END IF;
					
			-- Add the record to the list of user punishments
			INSERT INTO USER_PUNISHMENT (	ID_USER,
													ID_PUNISHER,
													ID_PUNISHMENT_TYPE,
													EVENT_DATE,
													END_DATE,
													REASON,
													ADMIN_NOTES)
			VALUES (	UserID,
					PunisherID,
					@ForcedPunishmentTypeID,
					NOW(),
					@PunishmentEndDate,
					Reason,
					AdminNotes);

		END IF;
		
		IF  @PunishmentBaseType = 2 OR @AddBan = 1 THEN
		
			UPDATE USER_INFO 
			SET 	ID_BAN_PUNISHMENT = (SELECT LAST_INSERT_ID()),
					BANNED = 1
			WHERE ID_USER = UserID;
			
		END IF;
		
		SELECT 	PunishmentTypeID, 
					@PunishmentBaseType PunishmentBaseType, 
					@PunishmentEndDate AS PunishmentEndDate, 
					@PunishmentSeverity AS PunishmentSeverity,
					@ForcedPunishmentTypeID AS ForcedPunishmentTypeID,
					@ForcedPunishmentBaseType AS ForcedPunishmentBaseType, 
					@ForcedPunishmentSeverity AS ForcedPunishmentSeverity,
					@AddBan AS AddBan,
					LAST_INSERT_ID() AS LastInsertedPunishment;
			
	END $$
	
DELIMITER ;


DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_GuildRolePermissionsUpsert`(
		IN `id_role_param` INT,
		IN `csv_string_param` TEXT
	)
	BEGIN
	-- Variables used for parsing the CSV string
	DECLARE comma_index INT;
	DECLARE permission_id  VARCHAR(10);
	
	-- Create a temporary table to store the parsed values from the CSV string
   CREATE TEMPORARY TABLE temp_permissions (id_permission INT);
	
	-- Loop to extract and insert values from the CSV string into the temporary table
	WHILE LENGTH(csv_string_param) > 0 DO
	  SET comma_index = INSTR(csv_string_param, ',');
	  IF comma_index = 0 THEN
	      SET permission_id = csv_string_param;
	      SET csv_string_param = '';
	  ELSE
	      SET permission_id = SUBSTRING(csv_string_param, 1, comma_index - 1);
	      SET csv_string_param = SUBSTRING(csv_string_param, comma_index + 1);
	  END IF;
	
	  -- Insert the parsed value into the temporary table
	  INSERT INTO temp_permissions (id_permission) VALUES (CAST(permission_id AS INT));
	END WHILE;
	
	
	DELETE FROM GUILD_ROLE_PERMISSION
	WHERE ID_ROLE = id_role_param;
		
	-- Insert the values from the temporary table into the main table
	INSERT INTO GUILD_ROLE_PERMISSION (ID_ROLE, ID_PERMISSION)
	SELECT id_role_param, id_permission FROM temp_permissions;
	
	-- Drop the temporary table
	DROP TEMPORARY TABLE IF EXISTS temp_permissions;
			
	END $$
	
DELIMITER ;


DELIMITER $$
	CREATE OR REPLACE PROCEDURE `sp_GuildRole_Delete`(
		IN `id_role_param` INT,
		IN `id_guild_param` INT
	)
	BEGIN
		
		DELETE FROM GUILD_ROLE_PERMISSION
		WHERE ID_ROLE = id_role_param;
		
		DELETE FROM GUILD_ROLE_ASSIGNED
		WHERE ID_ROLE = id_role_param
			AND ID_GUILD = id_guild_param;
	END $$

DELIMITER ;