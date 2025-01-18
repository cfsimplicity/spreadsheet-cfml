component{
    // Module Properties
    this.title = "Spreadsheet CFML";
    this.author = "Julian Halliwell";
    this.webURL = "https://github.com/cfsimplicity/spreadsheet-cfml";
    this.description = "CFML Spreadsheet Library";
    this.version = "4.3.0";
    this.autoMapModels = false;

    function configure(){
        binder.map( "Spreadsheet@spreadsheet-cfml" ).to( "#moduleMapping#.Spreadsheet" );
        binder.map( "Spreadsheet CFML" ).to( "#moduleMapping#.Spreadsheet" );
    }
}