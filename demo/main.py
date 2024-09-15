import json 
import requests
import xlwings as xw

from fastapi import Body, FastAPI, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import PlainTextResponse

app = FastAPI()


@app.post("/hello")
def hello(data: dict = Body):
    # Instantiate a Book object with the deserialized request body
    with xw.Book(json=data) as book:
        # Use xlwings as usual
        sheet = book.sheets[0]
        cell = sheet["A1"]
        if cell.value == "Hello xlwings!":
            cell.value = "Bye xlwings!"
        else:
            cell.value = "Hello xlwings!"

        # Pass the following back as the response
        return book.json()

@app.post("/apicall")
def api_call(data: dict = Body):

    headers = {
    'Authorization': 'Bearer ' + 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiI1NGY3ZDE3Ny05OWQ0LTQzNDktOTc1OC0zZTBkOGVkYWZkYWUiLCJlbWFpbCI6InRoaWVycnkubW91ZGlraS50ZWNodG9uaXF1ZUBnbWFpbC5jb20iLCJleHAiOjE3MjYzNjk5MTF9.Jt14vYeFrZjxEpvkfVkGBV9lK3TFNZ2q89d1k6fQakc',
    }

    params = {
        'method': 'RidgeCV',
        'n_hidden_features': '5',
        'lags': '25',
        'type_pi': 'gaussian',
        'h': '10',
    }

    files = {
        'file': open('/Users/t/Documents/datasets/time_series/univariate/USAccDeaths.csv', 'rb'),
    }

    response = requests.post('http://127.0.0.1:8000/forecastingreglinear', params=params, headers=headers, files=files)

    print(f"dir(response): {dir(response)}")

    #print(f"response.json(): {response.json()}")

    res = response.text
    print(f"res: {res}")
    print(f"dir(res): {dir(res)}")

    res2 = response.json()
    print(f"res2: {res2}")
    print(f"dir(res2): {dir(res2)}")

    # transform strings to lists
    mean = json.loads(res2["mean"])
    lower = json.loads(res2["lower"])
    upper = json.loads(res2["upper"])

    print(f"mean: {mean}")
    print(f"lower: {lower}")
    print(f"upper: {upper}")

    # Instantiate a Book object with the deserialized request body
    with xw.Book(json=data) as book:
        # Use xlwings as usual
        sheet = book.sheets[0]
        # Write the lists to Excel columns
        sheet.range("A1").value = "Mean"
        sheet.range("B1").value = "Lower"
        sheet.range("C1").value = "Upper"

        # Write the lists to Excel columns, one value per row
        for i in range(len(mean)):
            idx = i + 2
            print(f"mean[i]: {mean[i]}")
            sheet.range(f"A{idx}").value = mean[i]   # Write the 'mean' list in column A
            sheet.range(f"B{idx}").value = lower[i]  # Write the 'lower' list in column B
            sheet.range(f"C{idx}").value = upper[i]  # Write the 'upper' list in column C
        # Pass the following back as the response
        return book.json()

@app.exception_handler(Exception)
async def exception_handler(request, exception):
    # This handles all exceptions, so you may want to make this more restrictive
    return PlainTextResponse(
        str(exception), status_code=status.HTTP_500_INTERNAL_SERVER_ERROR
    )


# Office Scripts and custom functions in Excel on the web require CORS
cors_app = CORSMiddleware(
    app=app,
    allow_origins="*",
    allow_methods=["POST"],
    allow_headers=["*"],
    allow_credentials=True,
)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:cors_app", host="127.0.0.1", port=5000, reload=True)
