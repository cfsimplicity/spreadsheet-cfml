<cfscript>	
describe( "mergeCells", function(){

	beforeEach( function(){
		var data = querySim(
			"column1,column2
			a|b
			c|d");
		var xls = s.workbookFromQuery( data, false );
		var xlsx = s.workbookFromQuery( data=data, addHeaderRow=false, xmlFormat=true );
		variables.workbooks = [ xls, xlsx ];
	});

	it( "Retains merged cell data by default", function(){
		workbooks.Each( function( wb ){
			s.mergeCells( wb, 1, 2, 1, 2 );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "a" );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( "b" );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( "c" );
			expect( s.getCellValue( wb, 2, 2 ) ).toBe( "d" );
		});
	});

	it( "Can empty all but the top-left-most cell of a merged region", function(){
		workbooks.Each( function( wb ){
			s.mergeCells( wb, 1, 2, 1, 2, true )
				.write( wb, tempXlsPath, true );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "a" );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( "" );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( "" );
			expect( s.getCellValue( wb, 2, 2 ) ).toBe( "" );
		});
	});

	it( "Is chainable", function(){
		workbooks.Each( function( wb ){
			s.newChainable( wb ).mergeCells( 1, 2, 1, 2 );
			expect( s.getCellValue( wb, 1, 1 ) ).toBe( "a" );
			expect( s.getCellValue( wb, 1, 2 ) ).toBe( "b" );
			expect( s.getCellValue( wb, 2, 1 ) ).toBe( "c" );
			expect( s.getCellValue( wb, 2, 2 ) ).toBe( "d" );
		});
	});

	describe( "mergeCells throws an exception if", function(){

		it( "startRow OR startColumn is less than 1", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.mergeCells( wb, 0, 0, 1, 2 );
				}).toThrow( regex="Invalid" );
				expect( function(){
					s.mergeCells( wb, 1, 2, 0, 0 );
				}).toThrow( regex="Invalid" );
			});
		});

		it( "endRow/endColumn is less than startRow/startColumn", function(){
			workbooks.Each( function( wb ){
				expect( function(){
					s.mergeCells( wb, 2, 1, 1, 2 );
				}).toThrow( regex="Invalid" );
				expect( function(){
					s.mergeCells( wb, 1, 2, 2, 1 );
				}).toThrow( regex="Invalid" );
			});
		});

	});

	afterEach( function(){
		if( FileExists( variables.tempXlsPath ) ) FileDelete( variables.tempXlsPath );
		if( FileExists( variables.tempXlsxPath ) ) FileDelete( variables.tempXlsxPath );
	});
	
});	
</cfscript>