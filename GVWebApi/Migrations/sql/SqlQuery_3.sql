﻿ALTER TABLE [dbo].[PeriodHistoryViews] ALTER COLUMN [ActualVolume] [bigint] NULL
ALTER TABLE [dbo].[PeriodHistoryViews] ALTER COLUMN [VolumeOffset] [bigint] NULL
INSERT [dbo].[__MigrationHistory]([MigrationId], [ContextKey], [Model], [ProductVersion])
VALUES (N'201712100101514_third', N'GVWebapi.RemoteData.RevisionDBContext',  0x1F8B0800000000000400CD59DB6EE336107D2FD07F10F4D402592B97B6D806F22E123BC906DDC48695A4CFB434768852A42A52AEFD6D7DE827F5173AB4EE941C4BAE032CF2227338876786E41C8DF2EFDFFFB89FD721B356104B2AF8D03E1B9CDA16705F04942F8776A2161F3EDA9F3F7DFF9D7B13846BEB259F77A1E7A1279743FB55A9E8D271A4FF0A21918390FAB19062A106BE081D1208E7FCF4F457E7ECCC0184B011CBB2DC59C2150D61FB037F8E04F7215209610F220026B371B4785B54EB91842023E2C3D0BE7BF91DE624A283198442C19828625B578C12A4E2015BD816E15C28A290E8E5B3044FC5822FBD0807087BDA4480F3168449C802B82CA7778DE5F45CC7E2948E39949F4825C29E80671759721CD3FDA014DB45F2307D379866B5D1516F5338B4A71053117CA1B850BC79A1F0976D99AB5E8E58AC3D2AA94E7765D0703EB1F22927C5D9C023A4FF4EAC51C25412C39043A262C24EAC693267D4FF0D364FE20FE0439E3056258B74D1561BC0A1692C2288D566068B2C8419ACA85E49EFFCFDD8B69C3A86638214103BFCD358EFB9FAE527DB7A445264CEA0382395BC781836DC0187982808A6442988718BEF03D866B9C1C458F766367D0074B98B45128D41FAF9D27842F1B6D9D603597F05BE54AF431B1F6DEB96AE21C847323ACF9CE2E544271527B06F45BC5698791FC9BE08968460C4FAB6F395AFEFE3018EA3E9748609CA7DC6E0D39030DB9AC6F89495998FB6E5F944A7F97C1F5CCA60B2584850BD78A4A7755CA582CF4F34DC9FB71802AAE4B122B8BD9EE53B714DE4D1123362140FDE3B224FA7C702C4A39FF36CDCB8B73D67823181F2742C2613C4224BB85947C08F97B0DAD5AE457861F8BA4E59929B855A2789502C3066A9BAD61658AB966A8D1297156C9955853AD514D903B55B014A4AA984360B7D7B1005DD52AF9D54B07361777628BBFB40A2088B5E45E9B311CB4B657EF4C1EB2F7F618AE1F8B245050BB6C54A189D3E0B75AB2E9B01DCD2582AAD10F3EDBD1A0561635A737376243E5F6F57FE4DC92BB723F7D4CFA977CBDB4FCB6E198865766F31E010EFF6367628187660B545D1F781C47BF474A40B36EFA2D56FA1B6A86515B8C5DC1DBBA98B55E8A6B53B725D34ABA8754B0FAEB99AD628E683DD71EA325A05AB5BBA2356F5B58A571DEF1167AEB9B538F3C1EE380DB1ADE2358C3DF8B5A86D8D6A8BBD37BA56DC16503DDCEBE654C5D6B835555377CC52866BB7BB18ED8E640A7015CFB41D582D5AA2AE1B9BB8AE631448B3263B8DA26C741D66B57F4B2ECD29C5EA856C1AF2E86652B5BF3B6E68573A45BFD988150DB46E791BA9201C6CA5C3FB93A567AC9CF040385D8054699366A3B4FE6CF4D7DF4EAFEB4819B0BE0DEF37D372CEE992EACCEFED2B1B4DE9818D265F91D87F25F10F2159FF5885FD3FCD641E453FB4B6EEF23024A3DD0CF6BECAF7836F6B3F0F23DAEC47037C56DB7EB467C8F5FEF4D821EFE8578FBDCCEEFEF59D562AFBD9632FD0DADF1E764ACC86F7D854DB1BE07748485B43BC371FFDDAE366A3F62EAD6F2AA198A4B9C04052EA8D69F2E016B9A9EDAE53FD3EEE6209A7CB12427F2DE7E06BD12C41F339F77C21F2BDC0C0AB8CF229C656E13E112C44E42A5674810718CD3E48B9FD22F94258A277339C4370CF27898A12752525847356FBDCE93A6FAFBFFD0E50E7EC4E22FD4B1E2304A449752D9DF0EB84B2A0E07DDB54CB5D10FA2865128CAC3CA5A578B929901E05EF0894A56F0C78BFB4803F411831049313EE91151CC2ED59C25758127F93BFA2ED06D9BF11F5B4BB634A963109658651FAEBFFF938FA9F3E9FFE03DC38F191261A0000 , N'6.1.3-40302')
