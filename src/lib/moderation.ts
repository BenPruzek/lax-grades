
const BAD_WORDS = [
    "ass",
    "asshole",
    "bastard",
    "bitch",
    "bullshit",
    "cock",
    "cunt",
    "damn",
    "dick",
    "douche",
    "fag",
    "faggot",
    "fuck",
    "fucked",
    "fucking",
    "goddamn",
    "nigger",
    "nigga",
    "piss",
    "pussy",
    "shit",
    "shitty",
    "slut",
    "twat",
    "whore",
    "retard",
    "retarded",
    "crap",
    "dyke",
    "kike",
    "spic",
    "chink",
    "tranny",
    "shemale",
    "skank",
    "wanker",
    "bollocks",
    "bugger",
    "prick",
    "motherfucker",
    "tit",
    "tits",
    // Add more as needed
];

export function containsProfanity(text: string): boolean {
    if (!text) return false;
    const lowerText = text.toLowerCase();
    // Simple word boundary check
    return BAD_WORDS.some(word => {
        const regex = new RegExp(`\\b${word}\\b`, 'i');
        return regex.test(lowerText);
    });
}

export function moderateContent(text: string): { flagged: boolean; reason?: string } {
    if (containsProfanity(text)) {
        return { flagged: true, reason: "Content contains profanity." };
    }
    return { flagged: false };
}
