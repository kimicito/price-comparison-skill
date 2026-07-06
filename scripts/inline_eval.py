"""
Inline Eval — проверки во время работы (не пост-фактум).

Использование:
    from inline_eval import check_price_sanity, check_analog_brand, check_url_valid
    
    errors, warnings = inline_eval_item(item)
"""

from urllib.parse import urlparse


def check_price_sanity(price, reference_price=None, max_deviation=10.0):
    """Проверка, что цена адекватна.
    
    Args:
        price: проверяемая цена
        reference_price: эталонная цена (для сравнения)
        max_deviation: макс. отклонение в разах
    
    Returns: (ok, warning_msg)
    """
    if price is None:
        return True, None
    
    if not isinstance(price, (int, float)):
        return False, f"Цена не числовая: {price}"
    
    if price <= 0:
        return False, f"Цена <= 0: {price}"
    
    if reference_price and reference_price > 0:
        ratio = max(price, reference_price) / min(price, reference_price)
        if ratio > max_deviation:
            return False, f"Цена отличается более чем в {max_deviation}x: {price} vs {reference_price}"
        elif ratio > 3.0:
            return True, f"Цена отличается в {ratio:.1f}x — проверить"
    
    # Абсурдные цены
    if price < 100:
        return True, f"Цена подозрительно низкая: {price} ₽"
    if price > 10_000_000:
        return True, f"Цена подозрительно высокая: {price} ₽"
    
    return True, None


def check_url_valid(url):
    """Проверка, что URL корректный.
    
    Returns: (ok, error_msg)
    """
    if not url:
        return True, None  # URL опционален
    
    url = str(url).strip()
    if not (url.startswith('http://') or url.startswith('https://')):
        return False, f"URL не начинается с http:// или https://: {url}"
    
    try:
        parsed = urlparse(url)
        if not parsed.netloc:
            return False, f"URL без домена: {url}"
    except Exception as e:
        return False, f"URL некорректный: {url} ({e})"
    
    return True, None


def check_analog_brand(original_name, analog_name):
    """Проверка, что аналог — другой бренд (или та же марка с отклонениями).
    
    Returns: (ok, error_msg, is_same_brand)
    """
    if not original_name or not analog_name:
        return True, None, False
    
    original_brand = str(original_name).split()[0].lower()
    analog_brand = str(analog_name).split()[0].lower()
    
    is_same = original_brand == analog_brand
    
    if is_same:
        return True, f"Аналог той же марки ({analog_brand}) — отметить отклонения", True
    
    return True, None, False


def check_suppliers_different(supplier1, supplier2):
    """Проверка, что поставщики разные.
    
    Returns: (ok, warning_msg)
    """
    if not supplier1 or not supplier2:
        return True, "Только один поставщик — рекомендуется найти второго"
    
    s1 = str(supplier1).lower().strip()
    s2 = str(supplier2).lower().strip()
    
    if s1 == s2:
        return False, f"Поставщики одинаковые: {supplier1}"
    
    return True, None


def inline_eval_item(item):
    """Полная inline-проверка одной позиции.
    
    Args:
        item: dict с полями price1, price2, analog_price, url1, url2, analog_url, etc.
    
    Returns: (errors, warnings)
    """
    errors = []
    warnings = []
    
    # Проверка цен
    for field in ['price1', 'price2']:
        price = item.get(field)
        if price is not None:
            ok, msg = check_price_sanity(price)
            if not ok:
                errors.append(f"{field}: {msg}")
            elif msg:
                warnings.append(f"{field}: {msg}")
    
    # Проверка аналога
    analog_price = item.get('analog_price')
    if analog_price:
        ok, msg = check_price_sanity(analog_price, item.get('price1'))
        if not ok:
            errors.append(f"analog_price: {msg}")
        elif msg:
            warnings.append(f"analog_price: {msg}")
    
    # Проверка URL
    for field in ['url1', 'url2', 'analog_url']:
        url = item.get(field)
        ok, msg = check_url_valid(url)
        if not ok:
            errors.append(f"{field}: {msg}")
    
    # Проверка бренда аналога
    ok, msg, is_same = check_analog_brand(item.get('name'), item.get('analog'))
    if msg:
        if is_same:
            warnings.append(f"Аналог: {msg}")
        else:
            warnings.append(f"Аналог: {msg}")
    
    # Проверка поставщиков
    ok, msg = check_suppliers_different(item.get('supplier1'), item.get('supplier2'))
    if msg:
        if not ok:
            errors.append(f"Поставщики: {msg}")
        else:
            warnings.append(f"Поставщики: {msg}")
    
    return errors, warnings


def inline_eval_all(results):
    """Inline-проверка всех позиций.
    
    Returns: dict с суммарной статистикой
    """
    total_errors = 0
    total_warnings = 0
    items_with_errors = 0
    items_with_warnings = 0
    
    for item in results:
        errors, warnings = inline_eval_item(item)
        if errors:
            total_errors += len(errors)
            items_with_errors += 1
            print(f"  ❌ #{item.get('num', '?')}: {item.get('name', '')}")
            for e in errors:
                print(f"     ERROR: {e}")
        if warnings:
            total_warnings += len(warnings)
            items_with_warnings += 1
            if not errors:
                print(f"  ⚠️  #{item.get('num', '?')}: {item.get('name', '')}")
            for w in warnings:
                print(f"     WARN: {w}")
    
    print(f"\n📊 Inline Eval Summary:")
    print(f"   Позиций: {len(results)}")
    print(f"   С ошибками: {items_with_errors} ({total_errors} ошибок)")
    print(f"   С замечаниями: {items_with_warnings} ({total_warnings} предупреждений)")
    
    return {
        'total_items': len(results),
        'items_with_errors': items_with_errors,
        'items_with_warnings': items_with_warnings,
        'total_errors': total_errors,
        'total_warnings': total_warnings,
        'passed': total_errors == 0
    }
