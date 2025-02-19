<cfscript>
describe( "conditionalFormatting", ()=>{

	it( "can apply a format to cells only when a custom formula evaluates to true", ()=>{
		var formatting = s.newConditionalFormatting()
			.onCells( "A1:B1" )
			.when( "$A1<0" )
			.setFormat( { fontColor:"RED" } );
		spreadsheetTypes.Each( ( type )=>{
			var chainable = s.newChainable( type ).addConditionalFormatting( formatting );
			var appliedRules = formatting.rulesAppliedToCell( "B1" );
			expect( appliedRules ).toBeEmpty();
			chainable.setCellValue( -1, 1, 1 ); //set A1 to -1
			appliedColor = formatting.getFormatAppliedToCell( "B1" ).fontColor;
			expect( appliedColor ).toBe( "255,0,0" ); //RED
		})
		//alternate direct syntax
		spreadsheetTypes.Each( ( type )=>{
			//A1 value starts as 0. Rule will make A1 and B1 red if A1 is below zero.
			var wb = ( type == "xls" )? s.newXls(): s.newXlsx();
			s.addRow( wb, [ 0, 0 ] );
			formatting.addToWorkbook( wb );
			var appliedRules = formatting.rulesAppliedToCell( "B1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, -1, 1, 1 ); //set A1 to -1
			appliedColor = formatting.getFormatAppliedToCell( "B1" ).fontColor;
			expect( appliedColor ).toBe( "255,0,0" ); //RED
		})
	})

	it( "can apply a format to cells only when the cell value meets a condition", ()=>{
		spreadsheetTypes.Each( ( type )=>{
			//A1 and B1 values start as 0. Rule will make A1 red if A1 is LT zero.
			var wb = s.newChainable( type ).addRow( [ 0, 0 ] ).getWorkBook();
			var formatting = s.newConditionalFormatting()
				.onCells( "A1" )
				.whenCellValueIs( "LT", 0 )
				.setFormat( { fontColor:"RED" } )
				.addToWorkbook( wb );
			var appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, -1, 1, 1 ); //set A1 to -1
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
			appliedColor = formatting.getFormatAppliedToCell( "A1" ).fontColor;
			expect( appliedColor ).toBe( "255,0,0" );//RED
			// other comparison types
			//GT
			appliedRules = formatting.remove().whenCellValueIs( "GT", 0 ).addToWorkbook( wb ).rulesAppliedToCell( "A1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, 1, 1, 1 ); //set A1 to 1
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
			//LTE
			appliedRules = formatting.remove().whenCellValueIs( "LTE", 0 ).addToWorkbook( wb ).rulesAppliedToCell( "A1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, 0, 1, 1 ); //set A1 to 0
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
			//GTE
			appliedRules = formatting.remove().whenCellValueIs( "GTE", 1 ).addToWorkbook( wb ).rulesAppliedToCell( "A1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, 1, 1, 1 ); //set A1 to 1
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
			//EQ B1 value (zero)
			appliedRules = formatting.remove().whenCellValueIs( "EQ", "$B1" ).addToWorkbook( wb ).rulesAppliedToCell( "A1" );//B1 == 0
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, 0, 1, 1 ); //set A1 to 0
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
			//NEQ
			appliedRules =formatting.remove().whenCellValueIs( "NEQ", 0 ).addToWorkbook( wb ).rulesAppliedToCell( "A1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, 1, 1, 1 ); //set A1 to 1
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
			//BETWEEN
			appliedRules = formatting.remove().whenCellValueIs( "BETWEEN", 2, 3 ).addToWorkbook( wb ).rulesAppliedToCell( "A1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, 2, 1, 1 ); //set A1 to 2
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
			//NOT BETWEEN
			appliedRules = formatting.remove().whenCellValueIs( "NOT BETWEEN", 2, 3 ).addToWorkbook( wb ).rulesAppliedToCell( "A1" );
			expect( appliedRules ).toBeEmpty();
			s.setCellValue( wb, 1, 1, 1 ); //set A1 to 1
			appliedRules = formatting.rulesAppliedToCell( "A1" );
			expect( appliedRules ).toHaveLength( 1 );
		})
	})

	it( "can target a specific sheet name", ()=>{
		var formatting = s.newConditionalFormatting()
			.onCells( "A1:B1" )
			.onSheetName( "testSheet" )
			.when( "$A1<0" )
			.setFormat( { fontColor:"RED" } );
		spreadsheetTypes.Each( ( type )=>{
			//A1 value starts as 0. Rule will make A1 and B1 red if A1 is below zero.
			var chainable = s.newChainable( type )
				.createSheet( "testSheet" )
				.setActiveSheet( "testSheet" )
				.addRow( [ 0, 0 ] )
				.addConditionalFormatting( formatting );
			var appliedRules = formatting.rulesAppliedToCell( "B1" );
			expect( appliedRules ).toBeEmpty();
			chainable.setCellValue( -1, 1, 1 ); //set A1 to -1
			appliedColor = formatting.getFormatAppliedToCell( "B1" ).fontColor;
			expect( appliedColor ).toBe( "255,0,0" ); //RED
		})
	})

	it( "can target a specific sheet number", ()=>{
		var formatting = s.newConditionalFormatting()
			.onCells( "A1:B1" )
			.onSheetNumber( 2 )
			.when( "$A1<0" )
			.setFormat( { fontColor:"RED" } );
		spreadsheetTypes.Each( ( type )=>{
			//A1 value starts as 0. Rule will make A1 and B1 red if A1 is below zero.
			var chainable = s.newChainable( type )
				.createSheet( "testSheet" )
				.setActiveSheet( "testSheet" )
				.addRow( [ 0, 0 ] )
				.addConditionalFormatting( formatting );
			var appliedRules = formatting.rulesAppliedToCell( "B1" );
			expect( appliedRules ).toBeEmpty();
			chainable.setCellValue( -1, 1, 1 ); //set A1 to -1
			appliedColor = formatting.getFormatAppliedToCell( "B1" ).fontColor;
			expect( appliedColor ).toBe( "255,0,0" ); //RED
		})
	})

	it( "supports a range of font, border and fill pattern formats", ()=>{
		var format = {
			fontColor:"RED"
			,fontSize: 12
			,bold: true
			,italic: true
			,underline: "double"
			,bottomBorder: "thick"
			,bottomBorderColor: "0000FF"//blue
			,leftBorder: "thin"
			,leftBorderColor: "0,255,0"//green
			,rightBorder: "dashed"
			,rightBorderColor: "00FFFF"//cyan 0,255,255
			,topBorder: "dotted"
			,topBorderColor: "FF00FF"// 255,0,255
			,backgroundFillColor: "BLUE" // 0,0,51
			,foregroundFillColor: "RED" // 0,0,204
			,fillPattern: "diamonds"
		};
		var formatting = s.newConditionalFormatting()
			.onCells( "A1" )
			.whenCellValueIs( "EQ", 0 )
			.setFormat( format );
		var expected = {
			fontColor: "255,0,0"//red
			,fontSize: 12
			,bold: true
			,italic: true
			,underline: "double"
			,bottomBorder: "thick"
			,bottomBorderColor: "0,0,255"//blue
			,leftBorder: "thin"
			,leftBorderColor: "0,255,0"//green
			,rightBorder: "dashed"
			,rightBorderColor: "0,255,255"//cyan
			,topBorder: "dotted"
			,topBorderColor: "255,0,255"
			,backgroundFillColor: "0,0,255"//blue
			,foregroundFillColor: "255,0,0"//red
			,fillPattern: "diamonds"
		};
		spreadsheetTypes.Each( ( type )=>{
			s.newChainable( type ).addRow( [ 0 ] ).addConditionalFormatting( formatting );
			var actual = formatting.getFormatAppliedToCell( "A1" );
			expect( actual ).toBe( expected );
		})
	})

	describe( "throws an exception if", ()=>{

		it( "the comparison operator is invalid", ()=>{
			expect( ()=>{
				s.newConditionalFormatting()
					.onCells( "A1" )
					.whenCellValueIs( "INVALID OPERATOR", 0 )
					.setFormat( { color:"RED" } )
					.addToWorkbook( s.new() );
			}).toThrow( type="cfsimplicity.spreadsheet.invalidOperatorArgument" );
		})

		it( "a BETWEEN operator is used and formula2 is not supplied", ()=>{
			expect( ()=>{
				s.newConditionalFormatting()
					.onCells( "A1" )
					.whenCellValueIs( "BETWEEN", 0 )
					.setFormat( { color:"RED" } )
					.addToWorkbook( s.new() );
			}).toThrow( type="cfsimplicity.spreadsheet.missingSecondFormulaArgument" );
		})

	})

})	
</cfscript>