component accessors="true"{

	property name="libraryInstance";
	property name="rootPath";

	public any function init( required Spreadsheet libraryInstance ){
		this.setLibraryInstance( arguments.libraryInstance );
		this.setRootPath( GetDirectoryFromPath( GetCurrentTemplatePath() ) & "../" );
		return this;
	}

	public Spreadsheet function library(){
		return this.getLibraryInstance();
	}

	public any function getCellHelper(){
		return library().getCellHelper();
	}

	public any function getClassHelper(){
		return library().getClassHelper();
	}

	public any function getColorHelper(){
		return library().getColorHelper();
	}

	public any function getColumnHelper(){
		return library().getColumnHelper();
	}

	public any function getCommetHelper(){
		return library().getCommetHelper();
	}

	public any function getCsvHelper(){
		return library().getCsvHelper();
	}

	public any function getDataTypeHelper(){
		return library().getDataTypeHelper();
	}

	public any function getDateHelper(){
		return library().getDateHelper();
	}

	public any function getExceptionHelper(){
		return library().getExceptionHelper();
	}

	public any function getFileHelper(){
		return library().getFileHelper();
	}

	public any function getFontHelper(){
		return library().getFontHelper();
	}

	public any function getFormatHelper(){
		return library().getFormatHelper();
	}

	public any function getHeaderImageHelper(){
		return library().getHeaderImageHelper();
	}

	public any function getImageHelper(){
		return library().getImageHelper();
	}

	public any function getInfoHelper(){
		return library().getInfoHelper();
	}

	public any function getQueryHelper(){
		return library().getQueryHelper();
	}

	public any function getRangeHelper(){
		return library().getRangeHelper();
	}

	public any function getRowHelper(){
		return library().getRowHelper();
	}

	public any function getSheetHelper(){
		return library().getSheetHelper();
	}

	public any function getStringHelper(){
		return library().getStringHelper();
	}

	public any function getVisibilityHelper(){
		return library().getvisibilityHelper();
	}

	public any function getWorkbookHelper(){
		return library().getWorkbookHelper();
	}

}