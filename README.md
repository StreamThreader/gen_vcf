# gen_vcf
Set of scripts for convert excel xlsx to vcf vcard file

Script "gen_vcf_base64_photo.py photo.jpg" used for generate Base64 text from square jpeg image (contact avatar), compatible with excel limitation.
Then you can put this text to "Photo" cell in xlxx file.

File "user_list.xlsx" is special formated list of users.

Script "gen_vcf.py user_list.xlsx" used for generate single multi-vcard vcf file from xlsx file, compatible with Roundcube and Nextcloud.
You can change language field used from xlsx file, by set variable def_lang to EN, UK, RU.
