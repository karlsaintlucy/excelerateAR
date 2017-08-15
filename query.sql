SELECT row_number() over(ORDER BY i.number ASC),
    i.number AS invoice_num,
    'https://www.idealist.org/invoices/' || i.id AS invoice_link,
    li.description AS description,
    u.first_name || ' ' || u.last_name AS posted_by,
    o.name AS org_name,
    i.created::date AS posted_date,
    (i.created + INTERVAL '45 days')::date AS due_date,
    EXTRACT(EPOCH FROM(SELECT(NOW() -
        (i.created + INTERVAL '45 days')))/86400)::int AS days_overdue,
    li.unit_price AS amount_due
    FROM invoices AS i
    LEFT JOIN users AS u ON u.id = i.creator_id
    LEFT JOIN orgs AS o ON o.id = i.org_id
    LEFT JOIN line_items AS li ON li.invoice_id = i.id
    WHERE o.name = %s
    AND i.payment_settled = FALSE
    AND ((li.item_type = 'JOB') OR (li.item_type = 'INTERNSHIP'))
    AND EXTRACT(EPOCH FROM(SELECT(NOW() -
        (i.created + INTERVAL '45 days')))/86400)::int > 0
    --  https://stackoverflow.com/questions/3420982/opposite-of-inner-join
    AND NOT EXISTS(SELECT NULL
        FROM payment_parts AS pp
        WHERE li.id = pp.line_item_id)
    ORDER BY i.number ASC;