# A practical application for statistical analysis with Copilot in Excel           
## High-level overview                       
Understanding stock market trends is key to making informed investment decisions. Simple averages and trend lines can provide a starting point, but advanced statistical techniques like mean, median, standard deviation, and variance allow for a deeper understanding of stock performance and volatility. In this reading, you will explore how these techniques can be applied to stock data to uncover trends, assess risk, and make data-driven decisions.        

**Learning objectives**           
By the end of this reading, you will be able to:

use statistical techniques such as mean, median, standard deviation, and variance to analyze stock market data

interpret insights and trends within this example scenario

leverage Copilot in Excel to automate your statistical analysis process

**Why statistical techniques matter for stock market analysis**          
Stock market analysis goes beyond tracking prices over time. To fully understand a stock's performance and make informed investment decisions, you need to look into the variability, trends, and central tendencies of the stock. The following points describe why these techniques are crucial:

The mean provides the average price of a stock over a given period, offering a baseline for understanding the stock’s overall performance. For example, if the average price of a stock is $120, it gives investors a general sense of the stock's valuation.

The median reveals the middle price point in the dataset, providing a representative central value that isn't skewed by extreme highs or lows. However, it doesn't necessarily indicate where the majority of prices cluster.

Standard deviation measures the volatility of a stock, providing insights into the risk involved. A higher standard deviation indicates more price fluctuation, and therefore, higher risk.

Variance offers a deeper look at how spread out stock prices are over time, further helping investors assess the stock’s volatility and risk.

These techniques are invaluable for investors, helping them minimize risk and maximize returns by making data-driven decisions.

**Best practices for conducting stock market analysis**   
Use specific prompts: When analyzing stock data using Copilot, be as specific as possible with your prompts. For example, instead of asking Copilot to show stock performance, rather prompt it to Show the mean and median of [stock name] prices for Q1.

Reference correct fields: Use the exact field names in your prompts. If your dataset labels the stock price as "closing price", ensure that’s the term used in your commands.

Visualize data: Charts and graphs provide a clear visual representation of stock performance. For example, create a line graph to visualize trends, or a box and whisker plot to detect outliers in stock prices.

Refine and iterate: If your initial analysis doesn’t yield the insights you need, refine your prompts or add new variables to deepen your analysis. This is particularly useful when forecasting future stock trends based on multiple factors.

**Practical use cases for stock market analysis**       
Below are four practical use cases where these techniques can be applied to gain deeper insights into stock market data.

**1. Mean and median in price analysis**     
The mean and median are fundamental measures for understanding a stock’s overall performance over a given period. These techniques are especially useful when analyzing stock prices for trends, potential outliers, or consistency in performance.

Use case: Suppose you are tracking the stock prices of Company X over six months. If the mean price is $120 and the median is $115, this suggests the prices are relatively close and not skewed by extreme values. However, if the mean is $120 and the median is $90, it could indicate that a few high outliers have inflated the mean, and most prices are actually lower. This discrepancy may prompt further investigation into specific events or conditions that led to those outliers.

Actionable insight: Investors often use both the mean and median to assess whether the stock’s current price is reflective of its typical value. A significant difference between these values could signal periods of volatility or the presence of outliers.


The output provides a snapshot of the stock's central tendency, helping to gauge its general performance.

**2. Standard deviation in portfolio management**                 
Standard deviation is a critical tool for managing investment portfolios. It quantifies the volatility of a stock or a portfolio, offering insight into the level of risk an investor might be exposed to.

Use case: If you manage a diversified portfolio containing multiple stocks, calculating the standard deviation for each stock's returns will help you identify which assets carry more risk. For example, if Stock A has a standard deviation of 5% and Stock B has a standard deviation of 15%, Stock B is more volatile and carries higher risk. Investors who prefer lower-risk investments may choose to reduce exposure to high-standard-deviation stocks.

Actionable insight: A stock with a high standard deviation can offer higher returns, but it also comes with more risk. Investors often balance their portfolios by including both high-risk, high-reward stocks and more stable, low-volatility stocks.

Example prompt: What is the standard deviation of [stock name] returns over the past year?

This helps in understanding which stocks might bring volatility to your portfolio.

**3. Variance in identifying volatility**            
While variance is closely related to standard deviation, it provides an additional layer of analysis, focusing on the spread of stock prices over time. Investors often use variance to determine how much a stock's price deviates from the average, offering a detailed look at volatility.

Use case: Suppose you are analyzing the performance of two stocks, Stock RED and Stock GREEN. Over the past year, Stock RED has a variance of four, while Stock GREEN has a variance of 15. This means Stock GREEN has more price fluctuations compared to Stock RED. An investor focused on stability might prefer Stock RED, while a more aggressive investor looking for higher potential returns might accept the higher risk of Stock GREEN.

Actionable insight: High variance can indicate potential opportunities but also greater risk. By understanding variance, investors can make more informed decisions about whether to invest in or avoid certain stocks based on their volatility profiles.

Example prompt: Calculate the variance of [stock name] prices for the last quarter.

Copilot will assist you in understanding the level of consistency or fluctuation in a stock's performance.

**4. Using mean and standard deviation for risk assessment**         
Together, the mean and standard deviation are powerful tools for risk assessment. Investors often compare a stock's mean price with its standard deviation to understand how far prices deviate from the average, which can help in determining the risk associated with holding the stock.

Use case: Consider you are evaluating Stock Z, which has a mean price of $150 and a standard deviation of $30. This means most of the stock prices fluctuate between $120 and $180, indicating moderate volatility. However, if the standard deviation were $60, the prices would fluctuate more widely between $90 and $210, signalling higher volatility and risk. Investors can use this information to decide whether they are comfortable with the potential price swings.

Actionable insight: By comparing a stock’s mean to its standard deviation, investors can gauge how reliable the mean price is as a predictor of future performance. High deviation means higher risk, while lower deviation suggests more predictable performance.

Example prompt: Show the mean and standard deviation for [stock name] prices over the last year.

This gives you a clear picture of the stock's performance range and can help with evaluating the risks.

## Conclusion           
These practical use cases demonstrate how mean, median, standard deviation and variance can be applied to real-world stock data to make informed decisions. By understanding these metrics investors can assess volatility, uncover trends, and identify potential risks or opportunities. Copilot streamlines these calculations, enabling you to focus on strategic decision-making based on clear, data-driven insights.
