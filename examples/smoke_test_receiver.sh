#!/usr/bin/env bash
# Smoke test for the local mock receiver.
#
# Exercises every route and behaviour:
#   1. Health check
#   2. Validation handshake on each route
#   3. Notification dispatch on /file_uploaded (handler called for both items)
#   4. Notification dispatch on /item_reviewed (self-write filtered, one item processed)
#   5. Wrong clientState (handler not called)
#   6. Notification-level dedup (same notification ignored on second send)
#   7. Item-level dedup (different notification, same items ignored)
#
# Prerequisites:
#   uv run python examples/run_receiver_local.py   (in another terminal)
#
# Usage:
#   bash examples/smoke_test_receiver.sh

set -euo pipefail

BASE_URL="http://localhost:8000"
CLIENT_STATE="local-test-secret"
PASS=0
FAIL=0
TMPFILE=$(mktemp)
trap 'rm -f "$TMPFILE"' EXIT

# ── Helpers ──────────────────────────────────────────────────────────

check() {
    local description="$1"
    local expected_status="$2"
    local actual_status="$3"
    local body="$4"

    if [ "$actual_status" = "$expected_status" ]; then
        echo "  PASS  $description (HTTP $actual_status)"
        PASS=$((PASS + 1))
    else
        echo "  FAIL  $description (expected HTTP $expected_status, got $actual_status)"
        echo "        body: $body"
        FAIL=$((FAIL + 1))
    fi
}

check_body_contains() {
    local description="$1"
    local expected_substring="$2"
    local body="$3"

    if echo "$body" | grep -q "$expected_substring"; then
        echo "  PASS  $description (body contains '$expected_substring')"
        PASS=$((PASS + 1))
    else
        echo "  FAIL  $description (body missing '$expected_substring')"
        echo "        body: $body"
        FAIL=$((FAIL + 1))
    fi
}

# curl wrapper: writes body to TMPFILE, prints status code to stdout
do_get() {
    curl -s -o "$TMPFILE" -w "%{http_code}" "$1"
}

do_post() {
    local url="$1"
    shift
    curl -s -o "$TMPFILE" -w "%{http_code}" -X POST "$url" "$@"
}

post_notification() {
    local path="$1"
    local sub_id="${2:-sub-smoke-001}"
    local client_state="${3:-$CLIENT_STATE}"

    do_post "$BASE_URL$path" \
        -H "Content-Type: application/json" \
        -d "{
            \"value\": [{
                \"subscriptionId\": \"$sub_id\",
                \"changeType\": \"updated\",
                \"clientState\": \"$client_state\",
                \"resource\": \"sites/site-id/lists/list-id\",
                \"tenantId\": \"tenant-id-smoke\"
            }]
        }"
}

echo ""
echo "============================================================"
echo "  SMOKE TEST — Local Mock Receiver"
echo "============================================================"
echo ""
echo "  Target: $BASE_URL"
echo ""

# ── 1. Health check ──────────────────────────────────────────────────

echo "── 1. Health check ──"

STATUS=$(do_get "$BASE_URL/health")
BODY=$(cat "$TMPFILE")

check "GET /health returns 200" "200" "$STATUS" "$BODY"
check_body_contains "GET /health body has status=ok" '"status":"ok"' "$BODY"

echo ""

# ── 2. Validation handshake ──────────────────────────────────────────

echo "── 2. Validation handshake (per route) ──"

for ROUTE in "/file_uploaded" "/item_reviewed"; do
    TOKEN="smoke-token-$(echo "$ROUTE" | tr -d /)"
    STATUS=$(do_post "$BASE_URL${ROUTE}?validationToken=$TOKEN")
    BODY=$(cat "$TMPFILE")

    check "POST ${ROUTE}?validationToken echoes token" "200" "$STATUS" "$BODY"
    check_body_contains "POST ${ROUTE} body is the token" "$TOKEN" "$BODY"
done

echo ""

# ── 3. /file_uploaded — both items dispatched ────────────────────────

echo "── 3. /file_uploaded — handler called for both mock items ──"
echo "     (check server logs for two handle_new_file banners)"

STATUS=$(post_notification "/file_uploaded" "sub-file-001")
BODY=$(cat "$TMPFILE")

check "POST /file_uploaded returns 202" "202" "$STATUS" "$BODY"
check_body_contains "POST /file_uploaded body has status=accepted" '"status":"accepted"' "$BODY"

echo ""

# ── 4. /item_reviewed — self-write filtered ──────────────────────────

echo "── 4. /item_reviewed — self-write filtering ──"
echo "     (check server logs: only proc-002 processed, proc-001 filtered)"

STATUS=$(post_notification "/item_reviewed" "sub-review-001")
BODY=$(cat "$TMPFILE")

check "POST /item_reviewed returns 202" "202" "$STATUS" "$BODY"

echo ""

# ── 5. Wrong clientState ─────────────────────────────────────────────

echo "── 5. Wrong clientState — handler not called ──"
echo "     (check server logs: no handler banner for this request)"

STATUS=$(post_notification "/file_uploaded" "sub-bad-001" "wrong-secret")
BODY=$(cat "$TMPFILE")

check "POST /file_uploaded with wrong clientState returns 202" "202" "$STATUS" "$BODY"

echo ""

# ── 6. Notification-level dedup ──────────────────────────────────────

echo "── 6. Notification-level dedup — same notification sent twice ──"
echo "     (check server logs: handler should NOT fire again for sub-file-001)"

STATUS=$(post_notification "/file_uploaded" "sub-file-001")
BODY=$(cat "$TMPFILE")

check "POST /file_uploaded (duplicate) returns 202" "202" "$STATUS" "$BODY"

echo ""

# ── 7. Item-level dedup ──────────────────────────────────────────────

echo "── 7. Item-level dedup — new notification, same items ──"
echo "     (check server logs: handler should NOT fire — items already processed)"

STATUS=$(post_notification "/file_uploaded" "sub-file-002")
BODY=$(cat "$TMPFILE")

check "POST /file_uploaded (new notif, same items) returns 202" "202" "$STATUS" "$BODY"

echo ""

# ── Summary ──────────────────────────────────────────────────────────

echo "============================================================"
echo "  RESULTS: $PASS passed, $FAIL failed"
echo "============================================================"
echo ""
echo "  Also check the server terminal for handler log banners:"
echo "    - Test 3: Two 'handle_new_file' banners (file-001, file-002)"
echo "    - Test 4: One 'handle_human_review' banner (proc-002 only)"
echo "    - Test 5: No handler banner (wrong clientState)"
echo "    - Test 6: No handler banner (notification dedup)"
echo "    - Test 7: No handler banner (item-level dedup)"
echo ""

if [ "$FAIL" -gt 0 ]; then
    exit 1
fi
