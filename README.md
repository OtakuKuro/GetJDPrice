# Get JD Price
- 可以获取jd商品价格并输出到Excel表格中
- 修改商品货物id可改变查询的商品
- 随机id防止被验证码墙

```python
cpu = urllib2.urlopen('https://p.3.cn/prices/mgets?pduid='+ str(random.randint(100000,999999))+ '&skuIds=J_100004330867',timeout=5)
```

~~只是想买电脑每天自己查的麻烦才搞的~~
