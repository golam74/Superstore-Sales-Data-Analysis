import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns



# Read the dataset
df = pd.read_excel("C:/Users/MECH 5/OneDrive/Golam Israil/Supere_Sales12.xlsx", engine='openpyxl')


# Show top 5 rows
print(df.head())

# Basic info
print(df.info())

# Check for missing values
print(df.isnull().sum())

# Convert date columns to datetime
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])

# Check and remove duplicates
df.drop_duplicates(inplace=True)

print("Total Sales:", df['Sales'].sum())
print("Total Profit:", df['Profit'].sum())
print("Total Orders:", df['Order ID'].nunique())

df['Month'] = df['Order Date'].dt.to_period('M')
monthly_sales = df.groupby('Month')['Sales'].sum()

# Plot
monthly_sales.plot(kind='line', figsize=(10,5), title='Monthly Sales')
plt.xlabel("Month")
plt.ylabel("Sales")
plt.grid(True)
plt.show()

# Sales by Region / Category
# Sales by Region
region_sales = df.groupby('Region')['Sales'].sum().sort_values(ascending=False)

# Plot
region_sales.plot(kind='bar', title='Sales by Region')
plt.ylabel("Sales")
plt.show()


# Sales by Category
sns.barplot(x='Category', y='Sales', data=df, estimator=sum)
plt.title("Sales by Category")
plt.show()

#Top Products by Sales
top_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10)
top_products.plot(kind='barh', title='Top 10 Products by Sales')
plt.xlabel("Sales")
plt.show()

#Discount vs Profit Correlation
sns.scatterplot(x='Discount', y='Profit', data=df)
plt.title('Discount vs Profit')
plt.show()

#Category vs Sub-Category Sales (Stacked Plot)
pivot = df.pivot_table(index='Sub-Category', columns='Category', values='Sales', aggfunc='sum')
pivot.plot(kind='bar', stacked=True, figsize=(12,6))
plt.title("Category vs Sub-Category Sales")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.show()

#this is my program 

