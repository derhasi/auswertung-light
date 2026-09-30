DROP INDEX `fahrer_lizenz`;--> statement-breakpoint
-- Lizenzen, die zwischenzeitlich doppelt vergeben wurden, eindeutig machen (Suffix „-doppelt<ID>“ zum Nachbearbeiten)
UPDATE `fahrer` SET `lizenz` = `lizenz` || '-doppelt' || `id`
WHERE `lizenz` <> '' AND `id` NOT IN (SELECT MIN(`id`) FROM `fahrer` WHERE `lizenz` <> '' GROUP BY `lizenz`);--> statement-breakpoint
CREATE UNIQUE INDEX `fahrer_lizenz_eindeutig` ON `fahrer` (`lizenz`) WHERE `lizenz` <> '';
