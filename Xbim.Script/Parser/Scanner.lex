%using QUT.Xbim.Gppg;

%namespace Xbim.Script

%option verbose, summary, caseinsensitive, noPersistBuffer, out:Scanner.cs
%visibility internal

%{
	//all the user code is in XbimQueryScanerHelper

%}
 
%%

%{
		
%}
/* ************  skip white chars and line comments ************** */
"\t"	     {}
" "		     {}
[\n]		 {} 
[\r]         {} 
[\0]+		 {} 
\/\/[^\r\n]* {}   /*One line comment*/



/* ********************** Identifiers ************************** */
"$"[a-z$][a-z0-9_]*		        { return (int)SetValue(Tokens.IDENTIFIER); }

/* ********************** Operators ************************** */
"="	|
"equals" |
"is" |
"is equal to"					{  return ((int)Tokens.OP_EQ); }

"!=" |
"is not equal to" |
"is not" |
"does not equal" |
"doesn't equal" |
"doesn't"						{ return ((int)Tokens.OP_NEQ); }

">" |
"is greater than"				{  return ((int)Tokens.OP_GT); }

"<"	|
"is less than"					{  return ((int)Tokens.OP_LT); }

">=" |
"is greater than or equal to"	{  return ((int)Tokens.OP_GTE); }

"<=" |
"is less than or equal to"		{  return ((int)Tokens.OP_LTQ); }

"&&" |
"and"							{  return ((int)Tokens.OP_AND); }

"||" |
"or"							{  return ((int)Tokens.OP_OR); }

"~"	|
"contains" |
"is like"						{return ((int)Tokens.OP_CONTAINS);}

"!~" |
"does not contain" |
"doesn't contain" |
"is not like" |
"isn't like"					{return ((int)Tokens.OP_NOT_CONTAINS);}

";"		{  return (';'); }
","		{  return (','); }
":"		{  return (':'); }
"("		{  return ('('); }
")"		{  return (')'); }


"north from"		{return ((int)Tokens.NORTH_OF); }
"south from"		{return ((int)Tokens.SOUTH_OF); }
"west from"			{return ((int)Tokens.WEST_OF); }
"east from"			{return ((int)Tokens.EAST_OF); }
"above from"		{return ((int)Tokens.ABOVE); }
"below from"		{return ((int)Tokens.BELOW); }
"spatialy equal"	{return ((int)Tokens.SPATIALLY_EQUALS); }
"disjoint from"		{return ((int)Tokens.DISJOINT); }
"intersects with" |
"intersect with"	{return ((int)Tokens.INTERSECTS); }
"touches" |
"touch"				{return ((int)Tokens.TOUCHES); }
"crosses" |
"cross"				{return ((int)Tokens.CROSSES); }
"within"			{return ((int)Tokens.WITHIN); }
"overlaps" |
"overlap"			{return ((int)Tokens.OVERLAPS); }
"relates to" |
"relate to"		{return ((int)Tokens.RELATE); }

"the same"		{return (int)Tokens.THE_SAME; }
"deleted"		{return (int)Tokens.DELETED; }
"inserted"		{return (int)Tokens.INSERTED; }
"edited"		{return (int)Tokens.EDITED; }
					 
/* ********************** Keywords ************************** */
"select"			{ return (int)Tokens.SELECT;}
"set"			{ return (int)Tokens.SET;}
"for"			{ return (int)Tokens.FOR;}
"where"			{ return (int)Tokens.WHERE;}
"create"			{ return (int)Tokens.CREATE;}
"with name" |
"called"			{ return (int)Tokens.WITH_NAME; }
"description" |
"described as"			{ return (int)Tokens.DESCRIPTION ;} 
"new"			{ return (int)Tokens.NEW;}  /*is new*/								
"add"			{ return (int)Tokens.ADD;}
"to"			{ return (int)Tokens.TO; }
"as"			{ return (int)Tokens.AS; }
"remove"			{ return (int)Tokens.REMOVE; }
"from"			{ return (int)Tokens.FROM; }
"export" |
"dump"			{ return (int)Tokens.DUMP; }
"count"			{ return (int)Tokens.COUNT; }
"sum"			{ return (int)Tokens.SUM; }
"min" |
"minimal"		{ return (int)Tokens.MIN; }
"max" |
"maximal"		{ return (int)Tokens.MAX; }
"average"		{ return (int)Tokens.AVERAGE; }
"clear"			{ return (int)Tokens.CLEAR; }
"open"			{ return (int)Tokens.OPEN; }
"close"			{ return (int)Tokens.CLOSE; }
"save"			{ return (int)Tokens.SAVE; }
"in"			{ return (int)Tokens.IN; }
"it"			{ return (int)Tokens.IT; }
"every"			{ return (int)Tokens.EVERY; }
"validate"		{ return (int)Tokens.VALIDATE; }
"copy"			{ return (int)Tokens.COPY; }
"property set" |
"property_set" |
"propertyset"	{ return (int)Tokens.PROPERTY_SET; }

"name"			{ return (int)Tokens.NAME; }									
"predefinedtype" |
"predefined_type" |
"predefined type"			{ return (int)Tokens.PREDEFINED_TYPE; }
"type" |
"family"		{ return (int)Tokens.TYPE; }
"material"		{ return (int)SetValue(Tokens.MATERIAL); }
"thickness"		{ return (int)Tokens.THICKNESS; }  
"file"			{ return (int)Tokens.FILE; }
"cobie"			{ return (int)Tokens.COBIE; }
"model"			{ return (int)Tokens.MODEL; }
"summary"			{ return (int)Tokens.SUMMARY; }
"reference"			{ return (int)Tokens.REFERENCE; }
"classification"	{ return (int)Tokens.CLASSIFICATION; }
"group"			{ return (int)SetValue(Tokens.GROUP); }
"organization"			{ return (int)SetValue(Tokens.ORGANIZATION); }
"owner"			{ return (int)Tokens.OWNER; }
"layer set" |
"layer_set"	|
"layerset"		{ return (int)Tokens.LAYER_SET; }
"code"		{ return (int)Tokens.CODE; }
"rule" |
"rules"		{ return (int)Tokens.RULE; }
"property"		{ return (int)Tokens.PROPERTY; }
"attribute"		{ return (int)Tokens.ATTRIBUTE; }

"null" |
"undefined" |
"unknown"			{return (int)Tokens.NONDEF;}

"defined"			{return (int)Tokens.DEFINED;}

/* ********************     values        ****************** */
/* character groups in the \nnn format are in OCTAL! */
[\-\+]?[0-9]+	    {  return (int)SetValue(Tokens.INTEGER); }
[\-\+]?[0-9]*[\.][0-9]*	|
[\-\+\.0-9][\.0-9]+E[\-\+0-9][0-9]* { return (int)SetValue(Tokens.DOUBLE); }
[\"]([\n]|[\000\011-\041\043-\176\200-\377]|[\042][\042])*[\"]	{ return (int)SetValue(); }
[\']([\n]|[\000\011-\046\050-\176\200-\377]|[\047][\047])*[\']	{ return (int)SetValue(); }
".T." |
".F." |
true |
false	    { return (int)SetValue(Tokens.BOOLEAN); }
[a-z]+[a-z_\-0-9]*	{ return (int)ProcessString(); }


/* -----------------------  Epilog ------------------- */
%{
	yylloc = new LexLocation(tokLin,tokCol,tokELin,tokECol);
%}
/* --------------------------------------------------- */
%%


