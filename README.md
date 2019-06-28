# Replenishment optimization for VMI system

## Deploy in Google App Engine
gcloud app deploy app.yaml --project [PROJECT_NAME]

## Optimization Model

Dimensions: customer, product, truck, plant

Objectives: Replenish appropriate products and orders to customers within available supply at plants and truck constraint

Constraints:
- Supply available at plant
- Minimum inventory at customer
- Available plants for each customer
- Available plants for each truck
- Minimum volume at each truck
- Minimum price at each truck
- Minimum cost at each truck
- Maximum volume at each truck
- Maximum weight at each truck
- % Transportation cost at each truck
- Limit plant received at each truck

## Programming
- Code: [Python](https://www.python.org/)
- Optimization Code: [Google OR Tools](https://github.com/google/or-tools)
- Solver Engine: [CBC Coin-or branch and cut](https://github.com/coin-or/Cbc)
- Frontend: [Flask](http://flask.pocoo.org/)
