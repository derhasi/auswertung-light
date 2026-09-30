CREATE TABLE `fahrer_version` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`fahrer_id` integer NOT NULL,
	`lizenz` text NOT NULL,
	`klasse` text DEFAULT '' NOT NULL,
	`nachname` text DEFAULT '' NOT NULL,
	`vorname` text DEFAULT '' NOT NULL,
	`rookie_jahr` integer,
	`plz` text DEFAULT '' NOT NULL,
	`ort` text DEFAULT '' NOT NULL,
	`verein` text DEFAULT '' NOT NULL,
	`geburtsdatum` text DEFAULT '' NOT NULL,
	`alte_lizenz` text DEFAULT '' NOT NULL,
	`anlass` text DEFAULT '' NOT NULL,
	`erstellt_am` text NOT NULL
);
--> statement-breakpoint
CREATE INDEX `fahrer_version_fahrer` ON `fahrer_version` (`fahrer_id`);--> statement-breakpoint
CREATE TABLE `lauf_aenderung` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`start_id` integer NOT NULL,
	`nr` integer NOT NULL,
	`vorher` text,
	`nachher` text,
	`kommentar` text NOT NULL,
	`zeitpunkt` text NOT NULL,
	FOREIGN KEY (`start_id`) REFERENCES `start`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `lauf_aenderung_start` ON `lauf_aenderung` (`start_id`);--> statement-breakpoint
DROP INDEX `fahrer_lizenz_unique`;--> statement-breakpoint
CREATE INDEX `fahrer_lizenz` ON `fahrer` (`lizenz`);--> statement-breakpoint
ALTER TABLE `lauf` ADD `status` text DEFAULT 'ok' NOT NULL;--> statement-breakpoint
ALTER TABLE `lauf` ADD `kommentar` text;--> statement-breakpoint
ALTER TABLE `start` ADD `fahrer_id` integer;--> statement-breakpoint
ALTER TABLE `start` ADD `fahrer_version_id` integer;--> statement-breakpoint
-- Datenübernahme: erste Version für jeden vorhandenen Fahrer
INSERT INTO `fahrer_version` (`fahrer_id`, `lizenz`, `klasse`, `nachname`, `vorname`, `rookie_jahr`, `plz`, `ort`, `verein`, `geburtsdatum`, `alte_lizenz`, `anlass`, `erstellt_am`)
SELECT `id`, `lizenz`, `klasse`, `nachname`, `vorname`, `rookie_jahr`, `plz`, `ort`, `verein`, `geburtsdatum`, `alte_lizenz`, 'übernommen', `geaendert_am` FROM `fahrer`;--> statement-breakpoint
-- Vorhandene Nennungen über die Lizenz mit dem Fahrer und seiner Version verknüpfen
UPDATE `start` SET `fahrer_id` = (SELECT `f`.`id` FROM `fahrer` `f` WHERE `f`.`lizenz` = `start`.`lizenz` ORDER BY `f`.`id` LIMIT 1) WHERE `lizenz` <> '';--> statement-breakpoint
UPDATE `start` SET `fahrer_version_id` = (SELECT MAX(`v`.`id`) FROM `fahrer_version` `v` WHERE `v`.`fahrer_id` = `start`.`fahrer_id`) WHERE `fahrer_id` IS NOT NULL;
