component{
    // Module Properties
    this.title = "LuceeSpreadsheet";
    this.author = "Julian Halliwell";
    this.webURL = "https://github.com/cfsimplicity/lucee-spreadsheet";
    this.description = "Spreadsheet Library for Lucee";
    this.version = "2.0.1-develop";
    this.autoMapModels = false;

    function configure(){
        binder.map( "Spreadsheet@lucee-spreadsheet" ).to( "#moduleMapping#.Spreadsheet" );
        binder.map( "LuceeSpreadsheet" ).to( "#moduleMapping#.Spreadsheet" );
    }
}