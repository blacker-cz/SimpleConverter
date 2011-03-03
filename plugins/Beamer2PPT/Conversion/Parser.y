%using SimpleConverter.Contract

%namespace SimpleConverter.Plugin.Beamer2PPT

%start document

%{
    public Node Document { get; private set; }
%}

%union {
    public string Text;
    public Node documentNode;
    public List<Node> nodeList;
}

// todo: cleanup this hell :D
%token DOCUMENTCLASS "\documentclass", USEPACKAGE "\usepackage",
       TITLE "\title", AUTHOR "\author", TODAY "\today", DATE "\date", TITLEPAGE "\titlepage",
       BEGIN_DOCUMENT "\begin{document}", END_DOCUMENT "\end{document}",
       BEGIN_FRAME "\begin{frame}", END_FRAME "\end{frame}", FRAME "\frame", FRAMETITLE "\frametitle", PAUSE "\pause",
       BEGIN_ITEMIZE "\begin{itemize}", END_ITEMIZE "\end{itemize}", BEGIN_ENUMERATE "\begin{enumerate}",
       END_ENUMERATE "\end{enumerate}", BEGIN_DESCRIPTION "\begin{description}", END_DESCRIPTION "\end{description}",
       BEGIN_TABULAR "\begin{tabular}", END_TABULAR "\end{tabular}"
       SECTION "\section", SUBSECTION "\subsection", SUBSUBSECTION "\subsubsection",
       OVERLAY "overlay specification", OPTIONAL "optional parameter",
       TEXTBF "\textbf", TEXTIT "\textit", TEXTTT "\texttt", TEXTSC "\textsc",
        '{', '}', NL "\\ or \cr", '&', ENDROW "\\ or \cr",
       COLOR "\color", BFSERIES "\bfseries", TTFAMILY "\ttfamily", ITSHAPE "\itshape", SCSHAPE "\scshape",
       TINY "\tiny", SCRIPTSIZE "\scriptsize", FOOTNOTESIZE "\footnotesize", SMALL "\small",
       NORMALSIZE "\normalsize", LARGE "\large", LARGE2 "\Large", LARGE3 "\LARGE", HUGE "\huge", HUGE2 "\Huge",
       ITEM "\item"

%nonassoc <Text> STRING "plain text"

%right HIGH_PRIORITY

// todo: wtf??
%type <documentNode> command groupcommand standalonecommand commands slide titlesettings sectionsettings body
%type <nodeList> simpleformtext slidecontent bodycontent preambule

%%

document :
            documentclass preambule body    {
                                        Document = new Node("document");
                                        Document.Children = $2;
                                        Document.Children.Add($3);
                                    }
        ;

documentclass :
            DOCUMENTCLASS '{' STRING '}'            {
                                                        if(String.Compare($3, "beamer", false) != 0) {
                                                            Messenger.Instance.SendMessage("Invalid document class \"" + $3 + "\"", MessageLevel.ERROR);
                                                        }
                                                    }
        ;

preambule :                         {
                                        $$ = new List<Node>();
                                    }
        |   preambule USEPACKAGE '{' STRING '}'     {   // really need to process??
                                        Node tmp = new Node("package");
                                        tmp.Content = (object) $4;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   preambule titlesettings     {
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        ;

titlesettings :
            TITLE '{' simpleformtext '}'    {
                                        $$ = new Node("title");
                                        $$.Children = $3;
                                    }
        |   AUTHOR '{' simpleformtext '}'   {
                                        $$ = new Node("author");
                                        $$.Children = $3;
                                    }
        |   DATE '{' simpleformtext '}' {
                                        $$ = new Node("date");
                                        $$.Children = $3;
                                    }
        ;

sectionsettings :
            SECTION '{' simpleformtext '}'      {
                                        $$ = new Node("section");
                                        $$.Children = $3;
                                    }
        |   SUBSECTION '{' simpleformtext '}'   {
                                        $$ = new Node("subsection");
                                        $$.Children = $3;
                                    }
        |   SUBSUBSECTION '{' simpleformtext '}'    {
                                        $$ = new Node("subsubsection");
                                        $$.Children = $3;
                                    }
        ;

body : 
            BEGIN_DOCUMENT bodycontent END_DOCUMENT     {
                                        $$ = new Node("body");
                                        $$.Children = $2;
                                    }
        ;
    
bodycontent :                       {
                                        $$ = new List<Node>();
                                    }
        |   bodycontent titlesettings   {
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        |   bodycontent sectionsettings {
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        |   bodycontent slide       {
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        ;

slide : // todo: maybe count slides in here :)
            BEGIN_FRAME slidecontent END_FRAME   {
                                        $$ = new Node("slide");
                                        $$.Children = $2;
                                    }
        |   FRAME '{' slidecontent '}'   {
                                        $$ = new Node("slide");
                                        $$.Children = $3;
                                    }
        ;


slidecontent :                      {   /* return List<Node> - create node in specific command; append right side to the left side*/
                                        $$ = new List<Node>();
                                    }
        |   slidecontent '{' slidecontent '}'   {
                                        Node tmp = new Node("block");
                                        tmp.Children = $3;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   slidecontent STRING     {
                                        Node tmp = new Node("string");
                                        tmp.Content = (object) $2;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   slidecontent sectionsettings   {    // todo: insert to document, or create elsewhere on stack?
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        |   slidecontent environment/*    {
                                        $1.Add($2);
                                        $$ = $1;
                                    }*/
        |   slidecontent commands   {
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        ;

// todo: need to consider table environment
environment :
            BEGIN_ITEMIZE items_list END_ITEMIZE
        |   BEGIN_ENUMERATE items_list END_ENUMERATE
        |   BEGIN_DESCRIPTION items_list END_DESCRIPTION
        |   BEGIN_TABULAR '{' STRING '}' table_rows END_TABULAR
        ;

items_list :
            ITEM slidecontent
        |   items_list ITEM slidecontent
        ;

// todo: need to add rule for \hline
table_rows :
            table_cols
        |   table_rows ENDROW table_cols
        ;

// todo: need to add rule for \multicolumn
table_cols :
            slidecontent
        |   table_cols '&' slidecontent
        ;

commands : /* copy List<Node> from slidecontent to command Node*/
            command '{' slidecontent '}'    {
                                        $1.Children = $3;
                                        $$ = $1;
                                    }
        |   groupcommand slidecontent       {  // todo: resolve shift/reduce conflicts (dangling else???)
                                        $1.Children = $2;
                                        $$ = $1;
                                    }
        |   standalonecommand       {  // e.g. \today, \pause, \\
                                        $$ = $1;
                                    }
        ;

command :
            TEXTBF                  {
                                        $$ = new Node("bold");
                                    }
        |   TEXTIT                  {
                                        $$ = new Node("italic");
                                    }
        |   TEXTTT                  {
                                        $$ = new Node("typewriter");
                                    }
        |   TEXTSC                  {
                                        $$ = new Node("smallcaps");
                                    }
        ;

groupcommand :
            BFSERIES                {
                                        $$ = new Node("bold");
                                    }
        |   TTFAMILY                {
                                        $$ = new Node("typewriter");
                                    }
        |   ITSHAPE                 {
                                        $$ = new Node("italic");
                                    }
        |   SCSHAPE                 {
                                        $$ = new Node("smallcaps");
                                    }
        |   TINY                    {
                                        $$ = new Node("tiny");
                                    }
        |   SCRIPTSIZE              {
                                        $$ = new Node("scriptsize");
                                    }
        |   FOOTNOTESIZE            {
                                        $$ = new Node("footnotesize");
                                    }
        |   SMALL                   {
                                        $$ = new Node("small");
                                    }
        |   NORMALSIZE              {
                                        $$ = new Node("normalsize");
                                    }
        |   LARGE                   {
                                        $$ = new Node("large");
                                    }
        |   LARGE2                  {
                                        $$ = new Node("Large");
                                    }
        |   LARGE3                  {
                                        $$ = new Node("LARGE");
                                    }
        |   HUGE                    {
                                        $$ = new Node("huge");
                                    }
        |   HUGE2                   {
                                        $$ = new Node("Huge");
                                    }
        |   COLOR '{' STRING '}'    {
                                        $$ = new Node("color");
                                    }
        ;

standalonecommand :
            TODAY                   {
                                        $$ = new Node("today");
                                    }
        |   PAUSE                   {
                                        $$ = new Node("pause");
                                    }
        |   FRAMETITLE '{' simpleformtext '}'    {
                                        $$ = new Node("frametitle");
                                        $$.Children = $3;
                                    }
        |   NL                      {
                                        $$ = new Node("paragraph");
                                    }
        ;

// Simple formatted text; todo: check if this works how it should, resolve shift/reduce conflicts (dangling else??)
// ----------------------------------------------------------------------------

simpleformtext :                    {
                                        $$ = new List<Node>();
                                    }
        |   simpleformtext command '{' simpleformtext '}'   {
                                        $2.Children = $4;
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        |   simpleformtext groupcommand simpleformtext   {
                                        $2.Children = $3;
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        |   simpleformtext STRING   {
                                        Node tmp = new Node("string");
                                        tmp.Content = (object) $2;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   simpleformtext NL       {
                                        $1.Add(new Node("paragraph"));
                                        $$ = $1;
                                    }
        |   simpleformtext '{' simpleformtext '}'       {
                                        Node tmp = new Node("block");
                                        tmp.Children = $3;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   simpleformtext TODAY    {
                                        $1.Add(new Node("today"));
                                        $$ = $1;
                                    }
        ;

%%

public Parser(Scanner scn) : base(scn) { }
