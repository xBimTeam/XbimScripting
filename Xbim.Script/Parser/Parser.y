%{
	
%}
%namespace Xbim.Script
%partial   
%parsertype Parser
%output=Parser.cs
%visibility internal
%using Xbim.XbimExtensions.Interfaces
%using System.Linq.Expressions


%start expressions

%union{
		public string strVal;
		public int intVal;
		public double doubleVal;
		public bool boolVal;
		public Type typeVal;
		public IEnumerable<IPersistIfcEntity> entities;
		public object val;
	  }


%token	INTEGER	
%token	DOUBLE	
%token	STRING	
%token	BOOLEAN		
%token	NONDEF			/*not defined, null*/
%token	DEFINED			/*not null*/
%token	IDENTIFIER	
%token  OP_EQ			/*is equal, equals, is, =*/
%token  OP_NEQ			/*is not equal, is not, !=*/
%token  OP_GT			/*is greater than, >*/
%token  OP_LT			/*is lower than, <*/
%token  OP_GTE			/*is greater than or equal, >=*/
%token  OP_LTQ			/*is lower than or equal, <=*/
%token  OP_CONTAINS		/*contains, is like, */
%token  OP_NOT_CONTAINS	/*doesn't contain*/
%token  OP_AND
%token  OP_OR
%token  PRODUCT
%token  PRODUCT_TYPE
%token  FILE
%token  MODEL
%token  CLASSIFICATION
%token  PROPERTY_SET
%token  LAYER_SET
%token  REFERENCE
%token  ORGANIZATION
%token  OWNER
%token  CODE
%token  ATTRIBUTE
%token  PROPERTY
%token  COBIE
%token  SUMMARY
/*operations and keywords*/
%token  WHERE
%token  WITH_NAME /*with name, called*/
%token  DESCRIPTION /*and description, described as*/
%token  NEW /*is new*/
%token  ADD
%token  TO
%token  AS
%token  REMOVE
%token  FROM
%token  FOR
%token  NAME /*name*/
%token  PREDEFINED_TYPE
%token  TYPE
%token  MATERIAL
%token  THICKNESS
%token  GROUP
%token  IN
%token  IT
%token  EVERY
%token  COPY
%token  RULE

/* commands */
%token  SELECT
%token  SET
%token  CREATE
%token  DUMP
%token  CLEAR
%token  OPEN
%token  CLOSE
%token  SAVE
%token  COUNT
%token  SUM
%token  MIN
%token  MAX
%token  AVERAGE
%token  VALIDATE

/* spatial keywords */
%token  NORTH_OF
%token  SOUTH_OF
%token  WEST_OF
%token  EAST_OF
%token  ABOVE
%token  BELOW

%token  SPATIALLY_EQUALS
%token  DISJOINT
%token  INTERSECTS
%token  TOUCHES
%token  CROSSES
%token  WITHIN
%token  SPATIALLY_CONTAINS
%token  OVERLAPS
%token  RELATE

/* existance keywords */
%token THE_SAME	
%token DELETED	
%token INSERTED	
%token EDITED	

%%
expressions
	: expressions expression
	| expression
	;

expression
	: selection ';'
	| creation ';'
	| addition ';'
	| attr_setting ';'
	| variables_actions ';'
	| model_actions ';'
	| rule_check ';'
	| aggregation ';'
	;

attr_setting
	: SET value_setting_list FOR element_set			{EvaluateSetExpression($4.entities, ((List<Expression>)($2.val)));}
	;

value_setting_list
	: value_setting_list ',' value_setting				{((List<Expression>)($1.val)).Add((Expression)($3.val)); $$.val = $1.val;}
	| value_setting										{$$.val = new List<Expression>(){((Expression)($1.val))};}
	;

value_setting
	: attrOrProp TO value					{$$.val = GenerateSetExpression($1.strVal, $3.val, (Tokens)($1.val));}
	| MATERIAL TO IDENTIFIER				{$$.val = GenerateSetMaterialExpression($3.strVal);}
	;	

value
	: STRING								{$$.val = $1.strVal;}
	| BOOLEAN								{$$.val = $1.boolVal;}
	| INTEGER								{$$.val = $1.intVal;}
	| DOUBLE								{$$.val = $1.doubleVal;}
	| NONDEF								{$$.val = null;}
	;

num_value
	: DOUBLE								{$$.val = $1.doubleVal;}
	| INTEGER								{$$.val = $1.intVal;}
	;

model_actions
	: OPEN MODEL FROM FILE STRING																{OpenModel($5.strVal);}
	| CLOSE MODEL																				{CloseModel();}
	| VALIDATE MODEL																			{ValidateModel();}
	| SAVE MODEL TO FILE STRING																	{SaveModel($5.strVal);}
	| DUMP MODEL TO COBIE STRING STRING															{ExportCOBie($5.strVal,$6.strVal);}
	| DUMP MODEL TO FILE STRING	AS SUMMARY											            {ExportSummary($5.strVal);}
	| ADD REFERENCE MODEL STRING WHERE ORGANIZATION OP_EQ STRING OP_AND OWNER OP_EQ STRING		{AddReferenceModel($4.strVal, $8.strVal, $12.strVal);}
	/* | COPY IDENTIFIER TO MODEL STRING														{CopyToModel($2.strVal, $5.strVal);} */
	;

variables_actions
	: DUMP IDENTIFIER												{DumpIdentifier($2.strVal);}
	| CLEAR IDENTIFIER												{ClearIdentifier($2.strVal);}
	| DUMP string_list FROM element_set								{DumpAttributes($4.entities, ((List<string>)($2.val)), null, $4.strVal);}
	| DUMP string_list FROM element_set TO FILE STRING				{DumpAttributes($4.entities, ((List<string>)($2.val)), $7.strVal, $4.strVal);}
	;

aggregation
	: COUNT element_set												{$$.val = CountEntities($2.entities);}
	| SUM attrOrProp FROM element_set								{$$.val = SumEntities($2.strVal, (Tokens)($2.val), $4.entities);}
	| MIN attrOrProp FROM element_set								{$$.val = MinEntities($2.strVal, (Tokens)($2.val), $4.entities);}
	| MAX attrOrProp FROM element_set								{$$.val = MaxEntities($2.strVal, (Tokens)($2.val), $4.entities);}
	| AVERAGE attrOrProp FROM element_set							{$$.val = AverageEntities($2.strVal, (Tokens)($2.val), $4.entities);}
	;	

string_list
	: string_list ',' STRING										{((List<string>)($1.val)).Add($3.strVal); $$.val = $1.val;}
	| string_list ',' attribute										{((List<string>)($1.val)).Add($3.strVal); $$.val = $1.val;}
	| STRING														{$$.val = new List<string>(){$1.strVal};}
	| attribute														{$$.val = new List<string>(){$1.strVal};}
	;

selection
	: SELECT selection_statement									{Variables.Set("$$", ((IEnumerable<IPersistIfcEntity>)($2.val)));}
	| IDENTIFIER op_bool selection_statement						{AddOrRemoveFromSelection($1.strVal, ((Tokens)($2.val)), $3.val);}
	;

selection_statement
	: EVERY object														{$$.val = Select($2.typeVal);}
	| EVERY object STRING												{$$.val = Select($2.typeVal, $3.strVal);}
	| EVERY object WHERE conditions_set									{$$.val = Select($2.typeVal, ((Expression)($4.val)));}
	| EVERY CLASSIFICATION CODE STRING									{$$.val = SelectClassification($4.strVal);}
	;

element_set
	: IDENTIFIER													{$$.entities = GetVariableContent($1.strVal); $$.strVal = $1.strVal;}
	| selection_statement											{$$.entities = (IEnumerable<IPersistIfcEntity>)($1.val); }
	;
	
creation
	: CREATE creation_statement										{Variables.Set("$$", ((IPersistIfcEntity)($2.val)));}
	| CREATE CLASSIFICATION STRING									{CreateClassification($3.strVal);}
	| IDENTIFIER OP_EQ creation_statement							{Variables.Set($1.strVal, ((IPersistIfcEntity)($3.val)));}
	;

creation_statement
	: NEW object STRING												{$$.val = CreateObject($2.typeVal, $3.strVal);}			
	| NEW object WITH_NAME STRING 									{$$.val = CreateObject($2.typeVal, $4.strVal);}			
	| NEW object WITH_NAME STRING OP_AND DESCRIPTION STRING			{$$.val = CreateObject($2.typeVal, $4.strVal, $7.strVal);}			
	| NEW MATERIAL LAYER_SET STRING ':' layers						{$$.val = CreateLayerSet($4.strVal, (List<Layer>)($6.val));}			
	;

layers
	: layers ',' layer				{((List<Layer>)($1.val)).Add((Layer)($3.val)); $$.val = $1.val;}
	| layer							{$$.val = new List<Layer>(){(Layer)($1.val)};}
	;

layer
	: STRING num_value				{$$.val = new Layer(){material = $1.strVal, thickness = Convert.ToDouble($2.val)};}
	;

addition
	: ADD element_set TO IDENTIFIER									{AddOrRemove(Tokens.ADD, $2.entities, $4.strVal);}
	| REMOVE element_set FROM IDENTIFIER								{AddOrRemove(Tokens.REMOVE, $2.entities, $4.strVal);}
	;

conditions_set
	: '(' conditions ')' OP_AND '(' conditions ')'	{$$.val = Expression.AndAlso(((Expression)($2.val)), ((Expression)($6.val)));}
	| '(' conditions ')' OP_AND  condition			{$$.val = Expression.AndAlso(((Expression)($2.val)), ((Expression)($5.val)));}
	| '(' conditions ')' OP_OR '(' conditions ')'	{$$.val = Expression.OrElse(((Expression)($2.val)), ((Expression)($6.val)));}
	| '(' conditions ')' OP_OR condition			{$$.val = Expression.OrElse(((Expression)($2.val)), ((Expression)($5.val)));}
	| '(' conditions ')'							{$$.val = $2.val;}
	|  conditions									{$$.val = $1.val;}
	;

conditions
	: conditions OP_AND condition					{$$.val = Expression.AndAlso(((Expression)($1.val)), ((Expression)($3.val)));}
	| conditions OP_OR condition					{$$.val = Expression.OrElse(((Expression)($1.val)), ((Expression)($3.val)));}
	| condition										{$$.val = $1.val;}
	;
	
condition
	: materialCondition						{$$.val = $1.val;}
	| typeCondition							{$$.val = $1.val;}
	| groupCondition						{$$.val = $1.val;}
	| spatialCondition						{$$.val = $1.val;}
	| modelCondition						{$$.val = $1.val;}
	| existanceCondition					{$$.val = $1.val;}
	| attrOrPropCondition					{$$.val = $1.val;}
	| classificationCondition				{$$.val = $1.val;}
	;

classificationCondition
	: CLASSIFICATION CODE op_bool STRING	{$$.val = GenerateClassificationCondition($4.strVal, (Tokens)($3.val));}
	| CLASSIFICATION CODE op_bool NONDEF	{$$.val = GenerateClassificationCondition(null, (Tokens)($3.val));}
	| CLASSIFICATION CODE OP_NEQ DEFINED	{$$.val = GenerateClassificationCondition(null, Tokens.OP_EQ);}
	| CLASSIFICATION CODE OP_EQ DEFINED		{$$.val = GenerateClassificationCondition(null, Tokens.OP_NEQ);}
	;
		
materialCondition	
	: MATERIAL op_bool STRING			{$$.val = GenerateMaterialCondition($3.strVal, ((Tokens)($2.val)));}
	| MATERIAL op_cont STRING			{$$.val = GenerateMaterialCondition($3.strVal, ((Tokens)($2.val)));}
	
	| THICKNESS op_num_rel num_value	{$$.val = GenerateThicknessCondition(Convert.ToDouble($3.val), ((Tokens)($2.val)));}
	;
	
typeCondition	
	: TYPE op_bool PRODUCT_TYPE			{$$.val = GenerateTypeObjectTypeCondition($3.typeVal, ((Tokens)($2.val)));}
	| TYPE op_bool STRING				{$$.val = GenerateTypeObjectNameCondition($3.strVal, ((Tokens)($2.val)));}
	| TYPE op_cont STRING				{$$.val = GenerateTypeObjectNameCondition($3.strVal, ((Tokens)($2.val)));}
	| TYPE op_bool NONDEF				{$$.val = GenerateTypeObjectTypeCondition(null, ((Tokens)($2.val)));}
	| TYPE OP_NEQ DEFINED				{$$.val = GenerateTypeObjectTypeCondition(null, Tokens.OP_EQ);}
	| TYPE OP_EQ DEFINED				{$$.val = GenerateTypeObjectTypeCondition(null, Tokens.OP_NEQ);}
	| TYPE attrOrPropCondition			{$$.val = GenerateTypeCondition((Expression)($2.val));}
	;

groupCondition
	: GROUP attrOrPropCondition			{$$.val = GenerateGroupCondition((Expression)($2.val));}
	;

property
	: PROPERTY STRING							{$$.strVal = $2.strVal;}
	| PROPERTY NAME								{$$.strVal = "Name";}
	| PROPERTY DESCRIPTION						{$$.strVal = "Description";}
	| PROPERTY PREDEFINED_TYPE					{$$.strVal = "PredefinedType";}
	| property IN PROPERTY_SET STRING			{$$.strVal = $4.strVal + "\n" + $1.strVal;}
	;

attribute
	: NAME							{$$.strVal = "Name";}
	| DESCRIPTION					{$$.strVal = "Description";}
	| PREDEFINED_TYPE				{$$.strVal = "PredefinedType";}
	| ATTRIBUTE STRING				{$$.strVal = $2.strVal;}
	| ATTRIBUTE NAME				{$$.strVal = "Name";}
	| ATTRIBUTE DESCRIPTION			{$$.strVal = "Description";}
	| ATTRIBUTE PREDEFINED_TYPE		{$$.strVal = "PredefinedType";}
	;	

attrOrProp
	: STRING		{$$.strVal = $1.strVal; $$.val = Tokens.STRING;}
	| property		{$$.strVal = $1.strVal; $$.val = Tokens.PROPERTY;}
	| attribute		{$$.strVal = $1.strVal; $$.val = Tokens.ATTRIBUTE;}
	;

attrOrPropCondition	
	: attrOrProp op_num_rel INTEGER			{$$.val = GenerateValueCondition($1.strVal, $3.intVal, ((Tokens)($2.val)), (Tokens)($1.val));}
	| attrOrProp op_num_rel DOUBLE			{$$.val = GenerateValueCondition($1.strVal, $3.doubleVal, ((Tokens)($2.val)), (Tokens)($1.val));}
	  
	| attrOrProp op_bool STRING				{$$.val = GenerateValueCondition($1.strVal, $3.strVal, ((Tokens)($2.val)), (Tokens)($1.val));}
	| attrOrProp op_cont STRING				{$$.val = GenerateValueCondition($1.strVal, $3.strVal, ((Tokens)($2.val)), (Tokens)($1.val));}
	  
	| attrOrProp op_bool BOOLEAN			{$$.val = GenerateValueCondition($1.strVal, $3.boolVal, ((Tokens)($2.val)), (Tokens)($1.val));}
    | attrOrProp op_bool NONDEF				{$$.val = GenerateValueCondition($1.strVal, null, ((Tokens)($2.val)), (Tokens)($1.val));}
    | attrOrProp OP_NEQ DEFINED				{$$.val = GenerateValueCondition($1.strVal, null, Tokens.OP_EQ, (Tokens)($1.val));}
    | attrOrProp OP_EQ DEFINED				{$$.val = GenerateValueCondition($1.strVal, null, Tokens.OP_NEQ, (Tokens)($1.val));}
	;

spatialCondition
	: IT op_bool op_spatial IDENTIFIER			{$$.val = GenerateSpatialCondition((Tokens)($2.val), (Tokens)($3.val), $4.strVal);}
	;

modelCondition
	: MODEL op_bool	STRING						{$$.val = GenerateModelCondition(Tokens.MODEL, (Tokens)($2.val), $3.strVal);}
	| MODEL OWNER op_bool STRING				{$$.val = GenerateModelCondition(Tokens.OWNER, (Tokens)($3.val), $4.strVal);}
	| MODEL ORGANIZATION op_bool STRING			{$$.val = GenerateModelCondition(Tokens.ORGANIZATION, (Tokens)($3.val), $4.strVal);}
	;

existanceCondition
	: IT OP_EQ op_existance IN MODEL STRING								{$$.val = GenerateExistanceCondition((Tokens)($3.val), $6.strVal); }
	| IT op_bool IN MODEL STRING OP_AND IT op_bool IN MODEL STRING		{$$.val = GenerateExistanceCondition((Tokens)($2.val), $5.strVal, (Tokens)($8.val), $11.strVal); }
	;

op_existance
	: THE_SAME				{ $$.val = Tokens.THE_SAME; }
	| DELETED				{ $$.val = Tokens.DELETED; }
	| INSERTED				{ $$.val = Tokens.INSERTED;}
	| EDITED				{ $$.val = Tokens.EDITED;}
	;

op_spatial
	: NORTH_OF				{$$.val = Tokens.NORTH_OF			;}
	| SOUTH_OF				{$$.val = Tokens.SOUTH_OF			;}
	| WEST_OF				{$$.val = Tokens.WEST_OF			;}
	| EAST_OF				{$$.val = Tokens.EAST_OF			;}
	| ABOVE					{$$.val = Tokens.ABOVE				;}
	| BELOW					{$$.val = Tokens.BELOW				;}
	| SPATIALLY_EQUALS		{$$.val = Tokens.SPATIALLY_EQUALS	;}
	| DISJOINT				{$$.val = Tokens.DISJOINT			;}
	| INTERSECTS			{$$.val = Tokens.INTERSECTS			;}
	| TOUCHES				{$$.val = Tokens.TOUCHES			;}
	| CROSSES				{$$.val = Tokens.CROSSES			;}
	| WITHIN				{$$.val = Tokens.WITHIN				;}
	| OP_CONTAINS			{$$.val = Tokens.SPATIALLY_CONTAINS	;}
	| OVERLAPS				{$$.val = Tokens.OVERLAPS			;}
	| RELATE				{$$.val = Tokens.RELATE				;}
	;

op_bool
	: OP_EQ			{$$.val = Tokens.OP_EQ;}
	| OP_NEQ		{$$.val = Tokens.OP_NEQ;}
	;
	
op_num_rel
	: OP_GT			{$$.val = Tokens.OP_GT;}
    | OP_LT			{$$.val = Tokens.OP_LT;}
    | OP_GTE		{$$.val = Tokens.OP_GTE;}
    | OP_LTQ		{$$.val = Tokens.OP_LTQ;}
	| OP_EQ			{$$.val = Tokens.OP_EQ;}
	| OP_NEQ		{$$.val = Tokens.OP_NEQ;}
	;
	
op_cont
	: OP_CONTAINS		{$$.val = Tokens.OP_CONTAINS;}
	| OP_NOT_CONTAINS	{$$.val = Tokens.OP_NOT_CONTAINS;}
	;

object
	: PRODUCT				{$$.typeVal = $1.typeVal;}
	| PRODUCT_TYPE			{$$.typeVal = $1.typeVal;}
	| MATERIAL				{$$.typeVal = $1.typeVal;}
	| GROUP					{$$.typeVal = $1.typeVal;}
	| ORGANIZATION			{$$.typeVal = $1.typeVal;}
	;

rule_check
	: RULE STRING ':' conditions_set FOR element_set					{ CheckRule($2.strVal, (Expression)($4.val), $6.entities); }
	| RULE STRING ':' aggregation op_num_rel num_value					{ CheckRule($2.strVal, Convert.ToDouble($4.val), (Tokens)($5.val), Convert.ToDouble($6.val)); }
	| CLEAR RULE														{ ClearRules(); }
	| SAVE RULE TO FILE STRING											{ SaveRules($5.strVal); }
	;
	
%%
