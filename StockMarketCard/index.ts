import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as AdaptiveCards from "adaptivecards";

interface QuoteDetails {
	Symbol: string,
	TradingDay: Date,
	Price: number,
	Open: Date,
	High: number,
	Low: number,
	Change: number,
	ChangePercent: number
	
}

export class StockMarketCard implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _context: ComponentFramework.Context<IInputs>;
	private notifyOutputChanged: () => void;
	private _container: HTMLDivElement;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this._container = container;
		this._context = context;

		let symbol = context.parameters.Symbol.raw || "MSFT";
		let apiKey =   "NY26TP3X7NR5ME7I";

		
		this.getCard = this.getCard.bind(this);
		this.createCard = this.createCard.bind(this);

		this.getStockInfo(symbol, apiKey);
	}

	private getStockInfo(symbol: string, apiKey: string) {
		// fetch("https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol="+symbol+"&apikey="+apiKey)
			// var  proxy = "https://cors-anywhere.herokuapp.com/" ;
		fetch('https://query2.finance.yahoo.com/v8/finance/chart/'+ symbol)
		
			.then((response) => {
				return response.json();
			})
			.then((quoteJson) => {
				console.log(quoteJson);

				// const Data=JSON.parse(quoteJson);

				// console.log(Data.chart.result[symbol]);	 

				this.createCard(quoteJson);

			
			});
	}


	private createCard(quoteJson: any) {

		// const Data=quoteJson;
		var chartPreviousCloseValue = parseFloat(quoteJson.chart.result["0"].meta.chartPreviousClose);
		var regularMarketPriceValue= parseFloat(quoteJson.chart.result["0"].meta.regularMarketPrice);
		
		var CalculateChangeValue = chartPreviousCloseValue - regularMarketPriceValue ; 
		var CalculateChangePercentage = (CalculateChangeValue / chartPreviousCloseValue) *100 ;

		let quoteDetails: QuoteDetails = {
			// Symbol: Data.chart.result.meta.symbol,

			Symbol:quoteJson.chart.result["0"].meta.symbol,
				

			TradingDay: new Date(quoteJson.chart.result["0"].meta.regularMarketTime *1000),                                         
			Price: parseFloat(quoteJson.chart.result["0"].meta.regularMarketPrice), 
			Open: new Date( quoteJson.chart.result["0"].meta.currentTradingPeriod.regular.start),
			High: parseFloat(quoteJson.chart.result["0"].indicators.quote["0"].high),
			Low: parseFloat(quoteJson.chart.result["0"].indicators.quote["0"].low),
			Change: CalculateChangeValue,
			ChangePercent: CalculateChangePercentage
		};
		
		console.log(quoteDetails);

		let card = this.getCard(quoteDetails);

		// Create an AdaptiveCard instance
		let adaptiveCard = new AdaptiveCards.AdaptiveCard();

		// Set its hostConfig property unless you want to use the default Host Config
		// Host Config defines the style and behavior of a card
		adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
			fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
		});

		// Parse the card payload
		adaptiveCard.parse(card);

		// Render the card to an HTML element:
		let renderedCard = adaptiveCard.render();

		// And finally insert it somewhere in your page:
		this._container.appendChild(renderedCard);
	}

	private getCard(quoteDetails: QuoteDetails) {
		let arrowSymbol = "▲";
		let changeColor = "Good";

		if(quoteDetails.ChangePercent < 0){
			arrowSymbol = "▼";
			changeColor = "Attention";
		}

		let changeText = arrowSymbol + " " + 
			quoteDetails.Change.toFixed(2).toString() + " "+ 
			"(" + quoteDetails.ChangePercent.toFixed(2).toString()+ "% )"; 

		let card = {
			"$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
			"type": "AdaptiveCard",
			"version": "1.0",
			"speak": "Microsoft stock is trading at $62.30 a share, which is down .32%",
			"body": [
				{
					"type": "Container",
					"items": [
						{
							"type": "TextBlock",
							"text": quoteDetails.Symbol,
							"size": "Medium",
							"isSubtle": true
						},
						{
							"type": "TextBlock",
							"text": quoteDetails.TradingDay,
							"isSubtle": true
						}
					]
				},
				{
					"type": "Container",
					"spacing": "None",
					"items": [
						{
							"type": "ColumnSet",
							"columns": [
								{
									"type": "Column",
									"width": "stretch",
									"items": [
										{
											"type": "TextBlock",
											"text": quoteDetails.Price.toFixed(2).toString(),
											"size": "ExtraLarge"
										},
										{
											"type": "TextBlock",
											"text": changeText,
											"size": "Small",
											"color": changeColor,
											"spacing": "None"
										}
									]
								},
								{
									"type": "Column",
									"width": "auto",
									"items": [
										{
											"type": "FactSet",
											"facts": [
												{
													"title": "Open",
													"value": quoteDetails.Open
												},
												{
													"title": "High",
													"value": quoteDetails.High.toFixed(2).toString()
												},
												{
													"title": "Low",
													"value": quoteDetails.Low.toFixed(2).toString()
												}
											]
										}
									]
								}
							]
						}
					]
				}
			]
		};

		return card;
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
		this._context = context;
		
	var value =	this._context.parameters.Symbol.raw	; 

		
		// storing the latest context from the control.
		 
		let symbol:string = this._context.parameters.Symbol.raw || '';
	
		this.getStockInfo(symbol ,"hhh");

	
		
		
	}

	private wait(ms: any){
		var start = new Date().getTime();
		var end = start;
		while(end < start + ms) {
		  end = new Date().getTime();
	   }
	 }
	
	
	
	

	




	/** 
	 * It is called by the framework prior to a control receiving new data.  
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}