COPY wc.shipping_methods (id, title, description) FROM stdin;
flat_rate	Фиксированная ставка	Lets you charge a fixed rate for shipping.
free_shipping	Бесплатная доставка	Free shipping for qualifying orders.
local_pickup	Самовывоз	Allow customers to pick up orders.
\.

COPY wc.shipping_zone_locations (id, zone_id, code, type) FROM stdin;
1	1	US	country
2	2	US:CA	state
3	3	EU	continent
4	3	CA	country
5	3	GB	country
\.

COPY wc.shipping_zone_methods (id, zone_id, method_id, method_title, method_description, enabled, settings) FROM stdin;
1	1	flat_rate	Стандартная доставка	\N	t	{"cost": "9.99"}
2	1	flat_rate	Экспресс-доставка	\N	t	{"cost": "19.99"}
3	1	free_shipping	Бесплатная доставка (заказы > 100 ₽)	\N	t	{"min_amount": "100"}
4	2	flat_rate	Стандартная (Москва)	\N	t	{"cost": "5.99"}
5	2	local_pickup	Самовывоз из магазина	\N	t	{"cost": "0"}
6	3	flat_rate	Международная доставка	\N	t	{"cost": "29.99"}
\.

COPY wc.shipping_zones (id, name, "order") FROM stdin;
1	Доставка по РФ	0
2	Москва	1
3	Международная	2
\.

COPY wc.tax_classes (slug, name) FROM stdin;
standard	Standard
reduced-rate	Reduced Rate
zero-rate	Zero Rate
\.

COPY wc.tax_rates (id, country, state, postcode, city, rate, name, priority, compound, shipping, "order", class) FROM stdin;
1	Россия	CA			7.2500	НДС (Москва)	1	f	t	0	standard
2	Россия	CA	94*	Краснодар	8.6250	НДС (Краснодар)	1	f	t	0	standard
3	Россия	NY			8.0000	НДС (Уфа)	1	f	t	0	standard
4	Россия	TX			6.2500	НДС (Екатеринбург)	1	f	t	0	standard
5	Россия	WA			6.5000	НДС (Тюмень)	1	f	t	0	standard
6	Россия	FL			6.0000	НДС (Флорида)	1	f	t	0	standard
7	Россия				5.0000	НДС (пониженный)	1	f	t	0	reduced-rate
\.
