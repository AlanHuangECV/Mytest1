# 
FROM python:3.11

# 
WORKDIR /code

# 
COPY ./requirements.txt /code/requirements.txt

# 
RUN pip3 install --no-cache-dir --upgrade -r /code/requirements.txt

# 
COPY ./test_Neo_API15.py /code/test_Neo_API15.py

# 
CMD ["python", "/code/test_Neo_API15.py"]

