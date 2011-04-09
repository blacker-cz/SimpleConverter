%using SimpleConverter.Contract

%namespace SimpleConverter.Plugin.Beamer2PPT

%start document

%{
    public Node Document { get; private set; }
    public int SlideCount { get; private set; }
    public List<SectionRecord> SectionTable { get; private set; }
    public Dictionary<int, FrametitleRecord> FrametitleTable { get; private set; }
%}

%union {
    public string Text;
    public Node documentNode;
    public List<Node> nodeList;
    public HashSet<Node> nodeSet;
}

// todo: cleanup this hell :D
%token DOCUMENTCLASS "\documentclass", USEPACKAGE "\usepackage", USETHEME "\usetheme",
       TITLE "\title", AUTHOR "\author", TODAY "\today", DATE "\date", TITLEPAGE "\titlepage",
       BEGIN_DOCUMENT "\begin{document}", END_DOCUMENT "\end{document}",
       BEGIN_FRAME "\begin{frame}", END_FRAME "\end{frame}", FRAME "\frame", FRAMETITLE "\frametitle",
       FRAMESUBTITLE "\framesubtitle", PAUSE "\pause",
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
       ITEM "\item", UNDERLINE "\underline", AND "\and", TEXTCOLOR "\textcolor", HLINE "\hline", CLINE "\cline",
       MULTICOLUMN "\multicolumn"

%nonassoc <Text> STRING "plain text"
%nonassoc <Text> OPTIONAL "optional parameter"
%nonassoc <Text> OVERLAY "overlay specification"

// setup types for some non-terminals
%type <documentNode> command groupcommand standalonecommand commands slide titlesettings body environment documentclass preambule
%type <nodeList> simpleformtext slidecontent bodycontent items_list table_rows table_cols table_line
%type <Text> optional overlay
%type <nodeSet> table_line

%%

document :
            documentclass preambule body    {
                                        Document = $1;
                                        Document.Children = new List<Node>();
                                        Document.Children.Add($2);
                                        Document.Children.Add($3);
                                    }
        ;

documentclass :
            DOCUMENTCLASS optional '{' STRING '}'    {
                                        if(String.Compare($4, "beamer", false) != 0) {
                                            Messenger.Instance.SendMessage("Invalid document class \"" + $4 + "\"", MessageLevel.WARNING);
                                        }
                                        $$ = new Node("document");
                                        $$.OptionalParams = $2;
                                    }
        ;

preambule :                         {
                                        $$ = new Node("preambule");
                                        $$.Children = new List<Node>();
                                    }
        |   preambule USEPACKAGE '{' STRING '}'     {   // really need to process??
                                        Node tmp = new Node("package");
                                        tmp.Content = $4 as object;
                                        $1.Children.Add(tmp);
                                        $$ = $1;
                                    }
        |   preambule optional USETHEME '{' STRING '}'     {
                                        Node tmp = new Node("theme");
                                        tmp.Content = $5 as object;
                                        tmp.OptionalParams = $2;
                                        $1.Children.Add(tmp);
                                        $$ = $1;
                                    }
        |   preambule titlesettings     {
                                        $1.Children.Add($2);
                                        $$ = $1;
                                    }
        ;

titlesettings :
            TITLE '{' simpleformtext '}'    {
                                        $$ = new Node("title");
                                        $$.Children = $3;
                                    }
        |   AUTHOR '{' simpleformtext '}'   { // todo: can contain \and command
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
                                        SectionTable.Add(new SectionRecord(SlideCount + 1, $3));
                                    }
        |   SUBSECTION '{' simpleformtext '}'   {
                                        SectionTable.Add(new SectionRecord(SlideCount + 1, $3, SectionType.SUBSECTION));
                                    }
        |   SUBSUBSECTION '{' simpleformtext '}'    {
                                        SectionTable.Add(new SectionRecord(SlideCount + 1, $3, SectionType.SUBSUBSECTION));
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
                                        $$ = $1;
                                    }
        |   bodycontent slide       {
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        ;

slide :
            BEGIN_FRAME slidecontent END_FRAME   {
                                        $$ = new Node("slide");
                                        $$.Children = $2;
                                        SlideCount++;
                                    }
        |   BEGIN_FRAME '{' simpleformtext '}' slidecontent END_FRAME   {
                                        $$ = new Node("slide");
                                        $$.Children = $5;
                                        SlideCount++;
                                        SetFrameTitle(SlideCount, $3);
                                    }
        |   BEGIN_FRAME '{' simpleformtext '}' '{' simpleformtext '}' slidecontent END_FRAME   {
                                        $$ = new Node("slide");
                                        $$.Children = $8;
                                        SlideCount++;
                                        SetFrameTitle(SlideCount, $3);
                                        SetFrameSubtitle(SlideCount, $6);
                                    }
        |   FRAME '{' slidecontent '}'   {
                                        $$ = new Node("slide");
                                        $$.Children = $3;
                                        SlideCount++;
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
                                        tmp.Content = $2 as object;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   slidecontent sectionsettings {
                                        $$ = $1;
                                    }
        |   slidecontent environment    {
                                        $1.Add($2);
                                        $$ = $1;
                                    }
        |   slidecontent commands   {
                                        if($2 != null) {    // need to check because of frametitle and framesubtitle commands
                                            $1.Add($2);
                                        }
                                        $$ = $1;
                                    }
        ;

// todo: need to consider table environment
environment :
            BEGIN_ITEMIZE items_list END_ITEMIZE    {
                                        $$ = new Node("bulletlist");
                                        $$.Children = $2;
                                    }
        |   BEGIN_ENUMERATE items_list END_ENUMERATE    {
                                        $$ = new Node("numberedlist");
                                        $$.Children = $2;
                                    }
        |   BEGIN_DESCRIPTION items_list END_DESCRIPTION    {
                                        $$ = new Node("descriptionlist");
                                        $$.Children = $2;
                                    }
        |   BEGIN_TABULAR STRING table_rows END_TABULAR    {
                                        $$ = new Node("table");
                                        $$.Children = $3;
                                        $$.Content = $2 as object;
                                    }
        ;

items_list :
            ITEM overlay optional slidecontent       {
                                        Node tmp = new Node("item");
                                        tmp.OverlaySpec = $2;
                                        tmp.OptionalParams = $3;
                                        tmp.Children = $4;
                                        $$ = new List<Node>();
                                        $$.Add(tmp);
                                    }
        |   items_list ITEM overlay optional slidecontent    {
                                        Node tmp = new Node("item");
                                        tmp.OverlaySpec = $3;
                                        tmp.OptionalParams = $4;
                                        tmp.Children = $5;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        ;

table_rows :
            table_cols              {
                                        Node tmp = new Node("tablerow");
                                        tmp.Children = $1;
                                        $$ = new List<Node>();
                                        $$.Add(tmp);
                                    }
        |   table_line table_cols  {
                                        Node tmp = new Node("tablerow");
                                        tmp.Children = $2;
                                        $$ = new List<Node>();
                                        $$.AddRange($1);
                                        $$.Add(tmp);
                                    }
        |   table_rows ENDROW table_cols    {
                                        Node tmp = new Node("tablerow");
                                        tmp.Children = $3;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   table_rows ENDROW table_line table_cols    {
                                        Node tmp = new Node("tablerow");
                                        tmp.Children = $4;
                                        $1.AddRange($3);
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        ;

table_line :
            HLINE                   {
                                        $$ = new HashSet<Node>();
                                        $$.Add(new Node("hline"));
                                    }
        |   CLINE '{' STRING '}'    {
                                        $$ = new HashSet<Node>();
                                        Node tmp = new Node("cline");
                                        tmp.Content = $3 as object;
                                        $$.Add(tmp);
                                    }
        |   table_line HLINE       {
                                        $$.Add(new Node("hline"));
                                    }
        |   table_line CLINE '{' STRING '}'    {
                                        Node tmp = new Node("cline");
                                        tmp.Content = $4 as object;
                                        $$.Add(tmp);
                                    }
        ;

table_cols :
            slidecontent            {
                                        Node tmp = new Node("tablecolumn");
                                        tmp.Children = $1;
                                        $$ = new List<Node>();
                                        $$.Add(tmp);
                                    }
        |   MULTICOLUMN '{' STRING '}' '{' STRING '}' '{' slidecontent '}' {
                                        Node tmp = new Node("tablecolumn_merged");
                                        tmp.Content = $3 as object;
                                        tmp.OptionalParams = $6;
                                        tmp.Children = $9;
                                        $$ = new List<Node>();
                                        $$.Add(tmp);
                                    }
        |   table_cols '&' slidecontent     {
                                        Node tmp = new Node("tablecolumn");
                                        tmp.Children = $3;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        |   table_cols '&' MULTICOLUMN '{' STRING '}' '{' STRING '}' '{' slidecontent '}'   {
                                        Node tmp = new Node("tablecolumn_merged");
                                        tmp.Content = $5 as object;
                                        tmp.OptionalParams = $8;
                                        tmp.Children = $11;
                                        $1.Add(tmp);
                                        $$ = $1;
                                    }
        ;

commands : /* copy List<Node> from slidecontent to command Node*/
            command '{' slidecontent '}'    {
                                        $1.Children = $3;
                                        $$ = $1;
                                    }
        |   groupcommand slidecontent       {
                                        $1.Children = $2;
                                        $$ = $1;
                                    }
        |   standalonecommand       {  // e.g. \today, \pause, \\
                                        $$ = $1;
                                    }
        ;

command :
            TEXTBF overlay optional {
                                        $$ = new Node("bold");
                                        $$.OverlaySpec = $2;
                                        $$.OptionalParams = $3;
                                    }
        |   TEXTIT overlay optional {
                                        $$ = new Node("italic");
                                        $$.OverlaySpec = $2;
                                        $$.OptionalParams = $3;
                                    }
        |   TEXTTT overlay optional {
                                        $$ = new Node("typewriter");
                                        $$.OverlaySpec = $2;
                                        $$.OptionalParams = $3;
                                    }
        |   TEXTSC overlay optional {
                                        $$ = new Node("smallcaps");
                                        $$.OverlaySpec = $2;
                                        $$.OptionalParams = $3;
                                    }
        |   UNDERLINE overlay optional  { // beamer actually doesn't support this but we do :D
                                        $$ = new Node("underline");
                                        $$.OverlaySpec = $2;
                                        $$.OptionalParams = $3;
                                    }

        |   TEXTCOLOR overlay optional '{' STRING '}'    {
                                        $$ = new Node("color");
                                        $$.OverlaySpec = $2;
                                        $$.OptionalParams = $3;
                                        $$.Content = $5 as object;
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
        |   COLOR overlay optional '{' STRING '}'    {
                                        $$ = new Node("color");
                                        $$.OverlaySpec = $2;
                                        $$.OptionalParams = $3;
                                        $$.Content = $5 as object;
                                    }
        ;

standalonecommand :
            TODAY                   {
                                        $$ = new Node("today");
                                    }
        |   PAUSE                   {
                                        $$ = new Node("pause");
                                    }
        |   FRAMETITLE overlay optional '{' simpleformtext '}'   {
                                        SetFrameTitle(SlideCount + 1, $5, $2);
                                        $$ = null;
                                    }
        |   FRAMESUBTITLE overlay optional '{' simpleformtext '}'    {
                                        SetFrameSubtitle(SlideCount + 1, $5, $2);
                                        $$ = null;
                                    }
        |   NL                      {
                                        $$ = new Node("paragraph");
                                    }
        ;

optional :                          {
                                        $$ = "";
                                    }
        |   OPTIONAL                {
                                        $$ = $1;
                                    }
        ;

overlay :                           {
                                        $$ = "";
                                    }
        |   OVERLAY                 {
                                        $$ = $1;
                                    }
        ;

// Simple formatted text
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
                                        tmp.Content = $2 as object;
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
        |   simpleformtext PAUSE    {
                                        $1.Add(new Node("pause"));
                                        $$ = $1;
                                    }
        |   simpleformtext AND      {
                                        $1.Add(new Node("and"));
                                        $$ = $1;
                                    }
        ;

%%

public Parser(Scanner scn) : base(scn) {
    SlideCount = 0;
    SectionTable = new List<SectionRecord>();
    FrametitleTable = new Dictionary<int, FrametitleRecord>();
}

/// <summary>
/// Set frame title
/// </summary>
/// <param name="slide">Slide number</param>
/// <param name="content">Frame title content</param>
private void SetFrameTitle(int slide, List<Node> content, string overlay = "") {
    if(content == null || content.Count == 0)
        return;
    if(FrametitleTable.ContainsKey(slide)) {    // key exist change value
        FrametitleTable[slide].Title = content;
        FrametitleTable[slide].TitleOverlay = overlay;
    } else {    // key doesn't exist create new record
        FrametitleTable.Add(slide, new FrametitleRecord(content, null));
        FrametitleTable[slide].TitleOverlay = overlay;
    }
}

/// <summary>
/// Set frame subtitle
/// </summary>
/// <param name="slide">Slide number</param>
/// <param name="content">Frame subtitle content</param>
private void SetFrameSubtitle(int slide, List<Node> content, string overlay = "") {
    if(content == null || content.Count == 0)
        return;
    if(FrametitleTable.ContainsKey(slide)) {    // key exist change value
        FrametitleTable[slide].Subtitle = content;
        FrametitleTable[slide].SubtitleOverlay = overlay;
    } else {    // key doesn't exist create new record
        FrametitleTable.Add(slide, new FrametitleRecord(null, content));
        FrametitleTable[slide].SubtitleOverlay = overlay;
    }
}
