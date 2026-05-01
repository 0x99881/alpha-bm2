from __future__ import annotations


def test_core_pages_render_on_empty_database(client):
    for path in ["/", "/scores", "/score-overview", "/members", "/wear", "/profit-calendar"]:
        response = client.get(path)
        assert response.status_code in {200, 302}


def test_api_and_error_routes_are_stable(client):
    assert client.get("/api/score-overview").status_code == 200
    assert client.get("/scores?date=not-a-date").status_code == 200
    assert client.get("/missing-route").status_code == 404
    assert client.get("/supabase-push").status_code == 405


def test_score_post_rejects_bad_date_without_redirect(client):
    response = client.post("/scores/save", data={"date": "bad-date"})

    assert response.status_code == 200
    assert b"bad-date" in response.data
