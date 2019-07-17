import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as AdaptiveCards from "adaptivecards";

interface QuoteDetails {
	Symbol: string,
	TradingDay: string,
	Price: number,
	Open: number,
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

		this.getCard = this.getCard.bind(this);
		this.createCard = this.createCard.bind(this);

		this.getStockInfo("MSFT");
	}

	private getStockInfo(symbol: string) {
		fetch("https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=" + symbol + "&apikey=demo")
			.then((response) => {
				return response.json();
			})
			.then((quoteJson) => {
				console.log(quoteJson);

				this.createCard(quoteJson);
			});
	}

	private createCard(quoteJson: any) {

		let quoteDetails: QuoteDetails = {
			Symbol: quoteJson["Global Quote"]["01. symbol"],
			TradingDay: quoteJson["Global Quote"]["07. latest trading day"],
			Price: parseFloat(quoteJson["Global Quote"]["05. price"]),
			Open: parseFloat(quoteJson["Global Quote"]["02. open"]),
			High: parseFloat(quoteJson["Global Quote"]["03. high"]),
			Low: parseFloat(quoteJson["Global Quote"]["04. low"]),
			Change: parseFloat(quoteJson["Global Quote"]["09. change"]),
			ChangePercent: parseFloat(quoteJson["Global Quote"]["10. change percent"]
				.toString().replace("%", ""))
		};

		let card = this.getCard(quoteDetails);

		// Create an AdaptiveCard instance
		var adaptiveCard = new AdaptiveCards.AdaptiveCard();

		// Set its hostConfig property unless you want to use the default Host Config
		// Host Config defines the style and behavior of a card
		adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
			fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
			// More host config options
		});

		// Parse the card payload
		adaptiveCard.parse(card);

		// Render the card to an HTML element:
		var renderedCard = adaptiveCard.render();

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
													"value": quoteDetails.Open.toFixed(2).toString()
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