

import { AwsClient } from "aws4fetch";
import { Command } from "commander";

const program = new Command();
    program
        .name('cloudflare-url-signer')
        .description('A simple CLI that takes in all the necessary information to output a signed url')
        .version('1.0.0')
        
        .requiredOption('-u, --unsignedurl <Cloudflare R2 URL>', 'Define the path to the data Cloudflare URL')
        .requiredOption('-k, --key <storage key>', 'R2 Storage Key')
        .requiredOption('-s, --secret <storage secret>', 'R2 Storage Secret')
        .option('-e, --expiration <expiration time of link in seconds>', 'Expiration time in seconds for link');
    program.parse(process.argv);
    const options = program.opts();

export async function getPresignedUrl(
    unsignedUrl: string,
    storageId: string,
    storageKey: string,
    linkExpirationInSeconds?: number

) {
	const client = new AwsClient({
		accessKeyId: storageId,
		secretAccessKey: storageKey,
		// service: "s3", // R2 is S3-compatible
		region: "auto" // Cloudflare R2 handles region automatically
	});

	// const r2Endpoint = `https://${cloudflareAccountId}.r2.cloudflarestorage.com`;

	try {
		// Generate a presigned URL for GET operation
		// const unsignedUrl = `${r2Endpoint}/${privateBucketName}/${objectKey}`;
		const url = new URL(unsignedUrl);

		// Specify a custom expiry for the presigned URL, in seconds
        if(linkExpirationInSeconds)
		    url.searchParams.set("X-Amz-Expires", linkExpirationInSeconds.toString());

		const signed = await client.sign(
			new Request(url, {
				method: "GET"
			}),
			{
				aws: { signQuery: true }
			}
		);

		// Caller can now use this URL to upload to that object.
		return signed.url;
	} catch (error) {
		console.error("Error generating presigned URL:", error);
		throw error;
	}
}

const signedUrl = await getPresignedUrl(options.unsignedurl, options.key, options.secret)

console.log(signedUrl);