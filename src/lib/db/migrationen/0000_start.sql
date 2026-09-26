CREATE TABLE `fahrer` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
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
	`geaendert_am` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `fahrer_lizenz_unique` ON `fahrer` (`lizenz`);--> statement-breakpoint
CREATE TABLE `klasse` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`veranstaltung_id` integer NOT NULL,
	`name` text NOT NULL,
	`kuerzel` text DEFAULT '' NOT NULL,
	`position` integer DEFAULT 0 NOT NULL,
	`in_mannschaft` integer DEFAULT true NOT NULL,
	FOREIGN KEY (`veranstaltung_id`) REFERENCES `veranstaltung`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `klasse_veranstaltung` ON `klasse` (`veranstaltung_id`);--> statement-breakpoint
CREATE TABLE `lauf` (
	`start_id` integer NOT NULL,
	`nr` integer NOT NULL,
	`fehler1` integer DEFAULT 0 NOT NULL,
	`fehler2` integer DEFAULT 0 NOT NULL,
	`zeit` real,
	`import_id` text,
	`geaendert_am` text NOT NULL,
	PRIMARY KEY(`start_id`, `nr`),
	FOREIGN KEY (`start_id`) REFERENCES `start`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE TABLE `start` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`veranstaltung_id` integer NOT NULL,
	`klasse_id` integer NOT NULL,
	`startnummer` integer NOT NULL,
	`lizenz` text DEFAULT '' NOT NULL,
	`nachname` text DEFAULT '' NOT NULL,
	`vorname` text DEFAULT '' NOT NULL,
	`verein` text DEFAULT '' NOT NULL,
	`plz` text DEFAULT '' NOT NULL,
	`ort` text DEFAULT '' NOT NULL,
	`rookie_jahr` integer,
	`ausser_wertung` integer DEFAULT false NOT NULL,
	FOREIGN KEY (`veranstaltung_id`) REFERENCES `veranstaltung`(`id`) ON UPDATE no action ON DELETE cascade,
	FOREIGN KEY (`klasse_id`) REFERENCES `klasse`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE UNIQUE INDEX `start_startnummer` ON `start` (`veranstaltung_id`,`startnummer`);--> statement-breakpoint
CREATE INDEX `start_klasse` ON `start` (`klasse_id`);--> statement-breakpoint
CREATE TABLE `veranstaltung` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`name` text NOT NULL,
	`datum` text NOT NULL,
	`ort` text DEFAULT '' NOT NULL,
	`ausrichter` text DEFAULT '' NOT NULL,
	`zp_id` text DEFAULT '' NOT NULL,
	`strafe1` real DEFAULT 2 NOT NULL,
	`strafe2` real DEFAULT 10 NOT NULL,
	`fehler1_name` text DEFAULT 'Pylonen' NOT NULL,
	`fehler2_name` text DEFAULT 'Tore' NOT NULL,
	`mannschaft_anzahl` integer DEFAULT 6 NOT NULL,
	`urkunden_plaetze` integer DEFAULT 3 NOT NULL,
	`logo_links` text DEFAULT '' NOT NULL,
	`logo_rechts` text DEFAULT '' NOT NULL,
	`zeitquelle` text DEFAULT '' NOT NULL,
	`erstellt_am` text NOT NULL
);
