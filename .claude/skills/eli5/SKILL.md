---
name: eli5
description: Explain a topic like I'm a 5 year old. Use when the user types /eli5 <topic> or asks for a dead-simple picture explainer of how something works.
---

# eli5

Explain like I'm someone who knows nothing about this topic, using a HTML artifact with big pictures and few words.

Topic: $ARGUMENTS

---

Vendored from the `eli5` plugin by Thariq Shihipar, published in
https://github.com/anthropics/claude-plugins-community (plugin manifest declares
MIT; the marketplace repository ships an Apache-2.0 LICENSE at its root).

Kept in-repo rather than installed as a plugin so it loads from disk on the
first session, with no install step to race against skill loading.
